import type { BaseLayoutProps } from 'fumadocs-ui/layouts/shared';
import { Globe } from 'lucide-react';
import { SITE } from '../../../shared/sites';
import { Logo } from '../../web/app/components/Logo';
import { appName, gitConfig } from './shared';

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: (
        <>
          <Logo height={18} />
          {appName}
        </>
      ),
      url: SITE,
    },
    links: [{ type: 'icon', icon: <Globe />, text: 'Website', label: 'Website', url: SITE, external: true }],
    githubUrl: `https://github.com/${gitConfig.user}/${gitConfig.repo}`,
  };
}
