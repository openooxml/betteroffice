import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image, ImageDraw

from private_ssim import align, compare


class PrivateSsimTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def pages(self, name, lines, size=(32, 32), color='white'):
        directory = self.root / name
        directory.mkdir()
        for index in range(len(lines)):
            image = Image.new('RGB', size, color)
            ImageDraw.Draw(image).rectangle((4, 4, 12, 12), fill='black')
            image.save(directory / f'page_{index + 1:04d}.png')
        (directory / 'result.json').write_text(json.dumps(dict(status='ok', dpi=150, pages=len(lines))))
        (directory / 'pages_text.json').write_text(json.dumps(lines))
        (directory / 'pages.json').write_text(json.dumps([
            dict(page=index + 1, lines=[[line_index * 12, line] for line_index, line in enumerate(page)])
            for index, page in enumerate(lines)]))
        return directory

    def score(self, reference, actual, **options):
        return compare(reference, actual, self.root / 'compare', jobs=1, **options)

    def test_missing_and_extra_pages_are_zero_penalized(self):
        one = self.pages('one', [['alpha beta gamma delta epsilon']])
        two = self.pages('two', [['alpha beta gamma delta epsilon'], ['zeta eta theta iota kappa']])
        for reference, actual in [(one, two), (two, one)]:
            report = self.score(reference, actual)
            self.assertEqual(report['common_page_ssim'], 1.0)
            self.assertEqual(report['penalized_ssim'], 0.5)
            self.assertEqual(report['pages'][1]['ssim'], 0.0)
        empty = self.pages('empty', [])
        self.assertEqual(self.score(one, empty)['penalized_ssim'], 0.0)

    def test_one_pixel_adjustment_and_larger_size_mismatch(self):
        lines = [['alpha beta gamma delta epsilon']]
        reference = self.pages('reference', lines)
        for size in [(31, 33), (33, 31), (34, 32)]:
            actual = self.pages(f'actual-{size[0]}', lines, size=size)
            report = self.score(reference, actual)
            self.assertEqual(report['penalized_ssim'], 0.0 if size[0] == 34 else 1.0)
            self.assertIsNotNone(report['pages'][0]['size_note'])

    def test_alignment_distinguishes_page_index_and_break_agreement(self):
        a, b, c = ['alpha beta gamma delta epsilon', 'zeta eta theta iota kappa', 'lambda mu nu xi omicron']
        reference = self.pages('reference', [[a], [b], [c]])
        actual = self.pages('actual', [[a, b], [c]])
        report = self.score(reference, actual)
        self.assertEqual([row['aligned_page'] for row in report['pages']], [1, 1, 2])
        self.assertEqual([row['break_agrees'] for row in report['pages']], [True, False, True])
        self.assertEqual(report['first_divergence_page'], 2)
        self.assertEqual(report['page_agreement'], 1)
        self.assertEqual(report['page_agreement_fraction'], 1 / 3)
        self.assertEqual(report['break_agreement'], 2)
        self.assertEqual(report['aligned_pages'], 3)
        self.assertEqual(report['aligned_mean_ssim'], 1.0)

    def test_repeated_headers_and_numbered_footers_are_filtered(self):
        names = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta']
        body = [[f'{name} unique body words here'] for name in names]
        reference = self.pages('reference', [['Repeated running header', *page, f'Page {index + 1}']
                                                  for index, page in enumerate(body)])
        actual = self.pages('actual', body)
        report = self.score(reference, actual)
        self.assertEqual(report['page_agreement'], 6)
        self.assertEqual(report['break_agreement'], 6)
        self.assertIsNone(report['first_divergence_page'])

    def test_alignment_skips_unknown_words_and_only_searches_forward(self):
        anchor = 'alpha beta gamma delta epsilon'
        tail = 'zeta eta theta iota kappa'
        self.assertEqual(align([['unknown ' * 40 + anchor], [tail]], [[anchor, tail]]),
                         [(1, True), (1, False)])
        self.assertEqual(align([[anchor], [tail], [anchor]], [[anchor, tail]]),
                         [(1, True), (1, False), (None, False)])
        short = [['alpha beta gamma'], ['delta epsilon zeta eta theta']]
        self.assertEqual(align(short, short), [(1, True), (2, True)])

    def test_alignment_ignores_a_phrase_that_only_matches_much_later(self):
        title, toc = 'one two three four five', 'six seven eight nine ten'
        body = [' '.join(f'w{page}x{index}' for index in range(12)) for page in range(3)]
        reference = [[title], [toc], [body[0]], [body[1]], [body[2]]]
        actual = [[title], [body[0]], [body[1]], [body[2], toc]]
        self.assertEqual(align(reference, actual),
                         [(1, True), (None, False), (2, True), (3, True), (4, True)])

    def test_cache_reuse_and_baseline_regression(self):
        lines = [['alpha beta gamma delta epsilon'], ['zeta eta theta iota kappa']]
        reference = self.pages('reference', lines)
        actual = self.pages('actual', lines)
        cache = self.root / 'cache'
        baseline = self.score(reference, actual, cache=cache)
        baseline_path = self.root / 'baseline.json'
        baseline_path.write_text(json.dumps(baseline))
        with patch('private_ssim.ssim_pair', side_effect=AssertionError('cache missed')):
            unchanged = self.score(reference, actual, cache=cache, baseline=baseline_path)
        self.assertFalse(unchanged['regression'])
        self.assertEqual(unchanged['delta']['dropped_pages'], [])
        changed = self.pages('changed', [lines[0]], color='gray')
        report = self.score(reference, changed, cache=cache, baseline=baseline_path)
        self.assertTrue(report['regression'])
        self.assertEqual(report['delta']['actual_pages'], -1)
        self.assertEqual(report['delta']['page_agreement'], -1)
        self.assertEqual(report['delta']['first_divergence_page'], dict(before=None, after=2))
        self.assertEqual(report['delta']['dropped_pages'], [1, 2])
        self.assertEqual(json.loads((self.root / 'compare/score.json').read_text()), report)


if __name__ == '__main__':
    unittest.main()
