# Validate project mq-orghcart-hierarchy-client-side-solution

Date: 10/14/2024

## Findings

Following is the list of issues found in your project. [Summary](#Summary) of the recommended fixes is included at the end of the report.

### FN001035 @fluentui/react | Required

Install missing package @fluentui/react

Execute the following command:

```sh
npm i -SE @fluentui/react@^8.106.4
```

File: [./package.json](./package.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @fluentui/react@^8.106.4
npm dedupe
```