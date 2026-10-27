import type { Page } from 'playwright/test';

declare const evaluate: Page['evaluate'];
type PageFunction<Arg, R> = Exclude<Parameters<typeof evaluate<R, Arg>>[0], string>;

export async function pinnedEvaluate<R, Arg = void>(
  page: Page,
  pageFunction: PageFunction<Arg, R>,
  arg?: Arg,
): Promise<R> {
  const holder = await page.evaluateHandle<{
    run(arg: unknown): R | Promise<R>;
    promise?: R | Promise<R>;
  }>(`({ run: (${pageFunction}) })`);
  try {
    const result = await holder.evaluateHandle((holder, arg) => (holder.promise = holder.run(arg)), arg);
    try {
      return await result.jsonValue() as R;
    } finally {
      await result.dispose();
    }
  } finally {
    await holder.dispose();
  }
}
