import * as React from 'react';
import styles from './VendorSsoLauncher.module.scss';
import type { IVendorSsoLauncherProps } from './IVendorSsoLauncherProps';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SPHttpClient } from '@microsoft/sp-http';

interface IUserPayload {
  EnteredByEmail: string;
  EnteredBy: string;
  EnteredByLocation: string;
  exp?: number;
}

interface IUserProfileProperty {
  Key: string;
  Value: string;
}

const base64UrlEncodeBytes = (bytes: Uint8Array): string => {
  let binary: string = '';
  bytes.forEach((byte: number) => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
};

const base64UrlEncodeString = (text: string): string => {
  return base64UrlEncodeBytes(new TextEncoder().encode(text));
};

const generateSignedToken = async (
  payload: IUserPayload,
  expiresInSeconds: number,
  secret: string
): Promise<string> => {
  const fullPayload: IUserPayload = {
    ...payload,
    exp: Math.floor(Date.now() / 1000) + expiresInSeconds
  };

  const encodedPayload: string = base64UrlEncodeString(JSON.stringify(fullPayload));
  const key: CryptoKey = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );

  const signatureBuffer: ArrayBuffer = await crypto.subtle.sign(
    'HMAC',
    key,
    new TextEncoder().encode(encodedPayload)
  );

  return `${encodedPayload}.${base64UrlEncodeBytes(new Uint8Array(signatureBuffer))}`;
};

const getCurrentUserPayload = async (props: IVendorSsoLauncherProps): Promise<IUserPayload> => {
  const fallbackPayload: IUserPayload = {
    EnteredByEmail: props.userEmail || '',
    EnteredBy: props.userDisplayName || '',
    EnteredByLocation: 'Unknown Location'
  };

  try {
    const response = await props.spHttpClient.get(
      `${props.siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json'
        }
      }
    );

    if (!response.ok) {
      console.warn(`Profile request failed: ${response.status}. Using page context user data.`);
      return fallbackPayload;
    }

    const data = await response.json();
    const profile = data.d || data;
    const profileProperties = profile.UserProfileProperties;
    const allProps: IUserProfileProperty[] = Array.isArray(profileProperties)
      ? profileProperties
      : profileProperties?.results || [];
    const officeProp = allProps.find((property: IUserProfileProperty) => property.Key === 'Office');

    return {
      EnteredByEmail: profile.Email || fallbackPayload.EnteredByEmail,
      EnteredBy: profile.DisplayName || fallbackPayload.EnteredBy,
      EnteredByLocation: officeProp?.Value || fallbackPayload.EnteredByLocation
    };
  } catch (error) {
    console.warn('Profile request failed. Using page context user data.', error);
    return fallbackPayload;
  }
};

const VendorSsoLauncher: React.FC<IVendorSsoLauncherProps> = (props) => {
  const [isLaunching, setIsLaunching] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  const handleClick = async (): Promise<void> => {
    setErrorMessage('');

    if (!props.targetUrl) {
      setErrorMessage('Target URL is required.');
      return;
    }

    if (!props.sharedSecret) {
      setErrorMessage('Shared secret is required.');
      return;
    }

    setIsLaunching(true);

    try {
      const payload = await getCurrentUserPayload(props);
      const token = await generateSignedToken(
        payload,
        props.tokenExpirationSeconds || 300,
        props.sharedSecret
      );
      const launchUrl = `${props.targetUrl}${props.targetUrl.indexOf('?') === -1 ? '?' : '&'}token=${encodeURIComponent(token)}`;

      if (props.debugMode) {
        console.log('Customer complaint token payload:', payload);
        console.log('Customer complaint launch URL:', launchUrl);
        alert('Token generated. See browser console for details.');
        setIsLaunching(false);
        return;
      }

      window.location.href = launchUrl;
    } catch (error) {
      console.error('Customer complaint launch failed:', error);
      setErrorMessage('Unable to launch the application. Please try again.');
      setIsLaunching(false);
    }
  };

  return (
    <section className={`${styles.vendorSsoLauncher} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.container}>
        <PrimaryButton
          text={props.buttonLabel || 'Report a Customer Complaint'}
          onClick={handleClick}
          disabled={isLaunching}
          styles={{
            root: { backgroundColor: '#0078d4', border: 'none', color: '#ffffff', minWidth: '220px' },
            rootHovered: { backgroundColor: '#005a9e', color: '#ffffff' },
            rootPressed: { backgroundColor: '#004578', color: '#ffffff' },
            rootDisabled: { backgroundColor: '#c8c8c8', border: 'none', color: '#ffffff' }
          }}
        />
      </div>
      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error}>
          {errorMessage}
        </MessageBar>
      )}
    </section>
  );
};

export default VendorSsoLauncher;
