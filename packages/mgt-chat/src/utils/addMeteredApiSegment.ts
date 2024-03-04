export const addMeteredApiSegment = (useMeteredApi: boolean, url: string) => {
  // early exit if metered api is not enabled
  if (!useMeteredApi) {
    return url;
  }
  const urlHasExistingQueryParams = url.includes('?');
  return `${url}${urlHasExistingQueryParams ? '&' : '?'}model=B`;
};
