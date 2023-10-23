/**
 * @param value Entry needs to be checked.
 * @returns false if value is null or undefined
 */
export const isNull = (value: unknown) =>
  Object.is(value, null) || typeof value === "undefined";

/**
 * @param template The string was expected to contain a specific pattern.
 * @param pattern Regular expression to detect a template with syntax that needs to be replaced.
 * @param params An object containing values to replace matches the syntax in the template.
 * @returns An array of arrays with two elements e.g: ["{{name}}", "john doe"]
 */
export const detectReplacements = (
  template: string,
  pattern: RegExp,
  params: any
) => {
  const replaces: [string, string][] = [];
  const matches = template.match(new RegExp(pattern, "g"));
  if (!matches) {
    return replaces;
  }

  matches.forEach((match) => {
    const paramsKey = match.match(new RegExp(pattern))?.[1];
    if (!paramsKey) {
      return;
    }

    const replacement = params[paramsKey];
    if (isNull(replacement)) {
      return;
    }

    switch (typeof replacement) {
      case "string":
        replaces.push([match, replacement]);
        break;
    }
  });

  return replaces;
};
