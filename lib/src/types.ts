// import conf from "@azure/msal-common";
import { AccountInfo, AuthenticationResult, AuthError } from "@azure/msal-browser";
import { CacheOptions } from "@azure/msal-browser/dist/src/config/Configuration";
import * as msal from "@azure/msal-common";
import * as conf from "@azure/msal-common/dist/src/index";

// export type CacheOptions = conf.CacheOptions;
// export type AuthError = msal.AuthError;
// export type AuthResponse = msal.AuthResponse;
// export type SystemOptions = conf.SystemOptions;
// export type Account = msal.Account;

export type DataObject = {
  isAuthenticated: boolean;
  accessToken: string;
  idToken: string;
  user: User;
  custom: object;
  account?: AccountInfo;
};

export type FrameworkOptions = {
  globalMixin?: boolean;
};

export type Options = {
  auth: Auth;
  loginRequest: Request;
  tokenRequest: Request;
  resetRequest: ResetRequest;
  cache?: CacheOptions;
  framework?: FrameworkOptions;
};

export type Request = {
  scopes?: string[];
  account?: AccountInfo;
  state?: string;
};

export type ResetRequest = {
  authority?: string;
  scopes?: string[];
};
// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters,
// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
export type Auth = {
  clientId: string;
  authority: string;
  knownAuthorities: string[];
  redirectUri: string;
  autoRefreshToken?: boolean;
  onAuthentication: (
    ctx: object,
    error: AuthError,
    response: AuthenticationResult
  ) => any;
  onToken: (
    ctx: object,
    error: AuthError | null,
    response: AuthenticationResult | null
  ) => any;
  beforeSignOut: (ctx: object) => any;
};

export interface iMSAL {
  data: DataObject;
  signIn: () => Promise<any> | void;
  signOut: () => Promise<any> | void;
  acquireToken: () => Promise<any> | void;
  getTokenRedirect: () => Promise<any> | void;
  isAuthenticated: () => boolean;
}

export type User = {
  name: string;
  userName: string;
};
