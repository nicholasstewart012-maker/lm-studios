import { SPHttpClient } from '@microsoft/sp-http';

export interface IVendorSsoLauncherProps {
  sharedSecret: string;
  targetUrl: string;
  tokenExpirationSeconds: number;
  debugMode: boolean;
  buttonLabel: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  userDisplayName: string;
  userEmail: string;
  hasTeamsContext: boolean;
}
