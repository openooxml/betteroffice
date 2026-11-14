import { expect, test } from 'bun:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { LocaleProvider } from './LocaleContext';
import { useProposalTranslation } from './useProposalTranslation';

function ReviewLabel() {
  const { t } = useProposalTranslation();
  return <span>{t('proposals.layoutChanged')}</span>;
}

test('lazy review messages fall back to English without a locale', () => {
  expect(renderToStaticMarkup(<ReviewLabel />)).toBe('<span>Layout changed</span>');
  expect(renderToStaticMarkup(<LocaleProvider><ReviewLabel /></LocaleProvider>))
    .toBe('<span>Layout changed</span>');
});

test('lazy review messages prefer the supplied locale', () => {
  expect(renderToStaticMarkup(
    <LocaleProvider i18n={{ _lang: 'de', proposals: { layoutChanged: 'Layout geändert' } }}>
      <ReviewLabel />
    </LocaleProvider>
  )).toBe('<span>Layout geändert</span>');
});

test('null lazy review translations keep the English fallback', () => {
  expect(renderToStaticMarkup(
    <LocaleProvider i18n={{ _lang: 'de', proposals: { layoutChanged: null } }}>
      <ReviewLabel />
    </LocaleProvider>
  )).toBe('<span>Layout changed</span>');
});
