/** Allow only pinned jsDelivr font binaries fetched without credentials. */
export function isAllowedFontRequest(request) {
  const font =
    /^https:\/\/cdn\.jsdelivr\.net\/npm\/@betteroffice\/fonts(?:-cjk)?@(?:0\.1\.0|0\.2\.0)\/assets\/[A-Za-z0-9-]+\.(?:ttf|otf)$/.test(
      request.url,
    );
  return (
    font &&
    request.method === 'GET' &&
    request.bodyBytes === 0 &&
    !request.referer &&
    !request.cookie
  );
}
