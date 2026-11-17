import { RootProvider } from 'fumadocs-ui/provider/next';
import './global.css';
import type { Metadata } from 'next';
import '@fontsource-variable/inter';
import { appName, siteUrl } from '@/lib/shared';

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: {
    default: `${appName} documentation`,
    template: `%s | ${appName}`,
  },
  description:
    'Install and use BetterOffice editors and APIs for React, JavaScript, Rust, and Python.',
  openGraph: {
    type: 'website',
    siteName: `${appName} documentation`,
    url: siteUrl,
  },
  twitter: {
    card: 'summary_large_image',
  },
  robots: {
    index: true,
    follow: true,
  },
};

export default function Layout({ children }: LayoutProps<'/'>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body className="flex flex-col min-h-screen">
        <RootProvider>{children}</RootProvider>
      </body>
    </html>
  );
}
