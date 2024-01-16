'use client';

import Script from 'next/script';
import { useState } from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

export default function OfficeLayout({ children }: { children: React.ReactNode }) {
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);

  // Taskpane crash fix : Office.js removes history.pushstate
  // for more details :https://learn.microsoft.com/en-us/answers/questions/1150659/uncaught-typeerror-history-pushstate-is-not-a-func?orderby=helpful

  if (process.env.NEXT_PUBLIC_CLOUDFILES_ENV === 'development') {
    // @ts-ignore
    window._historyCache = {
      replaceState: window.history.replaceState,
      pushState: window.history.pushState,
    };
  }

  return (
    <>
      <Script
        src='https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js'
        onReady={() => {
          Office.onReady(() => setIsOfficeInitialized(true));
          if (process.env.NEXT_PUBLIC_CLOUDFILES_ENV === 'development') {
            // @ts-ignore
            window.history.replaceState = window._historyCache.replaceState;
            // @ts-ignore
            window.history.pushState = window._historyCache.pushState;
          }
        }}
      />
      <FluentProvider theme={webLightTheme}>{isOfficeInitialized ? children : <h1>Loading...</h1>}</FluentProvider>
    </>
  );
}
