# UK Address Finder PCF Control

React (virtual) Power Apps Component Framework control to search and select UK addresses via GetAddress.io.

## Features

- Virtual React control using shared platform React & Fluent UI libraries
- Debounced (300ms) address autocomplete
- GetAddress.io integration with caching & 5s timeout
- Manual entry fallback
- Outputs normalized address fields: addressLine1, addressLine2, city, county, postcode, country
- Accessible keyboard navigation (arrow keys, enter, escape)
- Inline highlighting of matched text, loading + empty states in dropdown

## Implementation details

The README above describes the user-visible features; the paragraphs below document exact implementation choices used in the control so maintainers and reviewers can quickly understand behavior and performance trade-offs.

- Provider and network

  - Uses GetAddress.io when a valid `apiKey` is provided (provider name: `GetAddress`).
  - Base API host: `https://api.getAddress.io`.
  - Autocomplete endpoint is called as `/autocomplete/{query}?api-key={key}&all=true&top=20` to request broader results (up to 20 suggestions).
  - Details endpoint is `/get/{id}?api-key={key}`.
  - All remote calls use a 5000ms (5s) timeout and are cancellable with AbortSignal.

- Caching

  - Results are cached in-memory using simple Map caches for both autocomplete (search) results and address details.
  - Cache TTL is 5 minutes (300000 ms). Cached entries are returned while fresh to reduce network calls.

- Search behaviour

  - Input is debounced by 300ms before issuing searches.
  - Minimum search length is 3 characters for most queries. Shorter queries that look like UK postcodes are allowed (the control applies a postcode pattern test to allow earlier postcode searches).
  - If the provider returns suggestions, the dropdown shows up to the returned suggestions (the code requests top=20).

- Keyboard & accessibility

  - Dropdown is rendered with `role="listbox"` and each suggestion uses `role="option"`.
  - `aria-live="polite"` is set on the results container to announce loading/no-result states.
  - Keyboard navigation supported: ArrowUp/ArrowDown to move, Enter to select, Escape to close.
  - Error messages are rendered with `role="alert"`.

- Manual entry fallback

  - If no provider is created (for example the `apiKey` is not set), the control falls back to a manual entry form.
  - Manual mode exposes fields: Address Line 1, Address Line 2, City/Town, County, Postcode, Country. Updates to these fields immediately populate the control outputs.

- Outputs / data contract

  - The control normalizes and exposes these output fields (matching the manifest):
    - `addressLine1`, `addressLine2`, `city`, `county`, `postcode`, `country`.
  - When a suggestion is selected the control will fetch full address details (if needed) and populate the above outputs; manual entries produce an `id` of `manual`.

- UI details
  - To avoid bringing the full Fluent icon bundles into the build, two small inline SVG components are used in-code for the search and close icons (named `SearchIcon` and `CloseIcon`).
  - Loading state shows a small Spinner and a short message; empty state and error states render user-friendly messages inside the dropdown area.

These details were taken directly from the implementation to keep documentation in sync with behaviour. If you change timeouts, debounce, caching or the API parameters, please update this section accordingly.

## Properties

| Name         | Usage  | Type            | Direction | Description                       |
| ------------ | ------ | --------------- | --------- | --------------------------------- |
| addressLine1 | output | SingleLine.Text | Output    | First address line                |
| addressLine2 | output | SingleLine.Text | Output    | Second address line               |
| city         | output | SingleLine.Text | Output    | Town/City                         |
| county       | output | SingleLine.Text | Output    | County                            |
| postcode     | output | SingleLine.Text | Output    | UK postcode                       |
| country      | output | SingleLine.Text | Output    | Country (United Kingdom)          |
| apiProvider  | input  | Enum            | Input     | Address API provider (GetAddress) |
| apiKey       | input  | SingleLine.Text | Input     | API key for provider              |
| placeholder  | input  | SingleLine.Text | Input     | Search box placeholder text       |

## Build

Run:

```
npm install
npm run refreshTypes
npm run build
```

## Usage Notes

- Provide a valid GetAddress.io API key bound or static for search to function; empty key forces manual entry mode.
- Minimum 3 characters required before a standard search is executed (UK postcode patterns are allowed sooner as they are typed – e.g. `EC1` will trigger after 3 chars, full postcode formats also work with/without space).
- Timeout after 5s triggers error state; user can retry or switch to manual entry.
- The manifest now declares an external service domain `api.getaddress.io`. Ensure your environment allows this premium control scenario and that your solution publisher approvals are in place.
- To customise placeholder text, set the `placeholder` input property.
- If you change the API key at runtime the provider is recreated automatically.

## Accessibility

- Listbox semantics, keyboard navigation (Up/Down/Enter/Escape)
- ARIA alert for error messages

## Next Improvements (Backlog)

- Full manual form with all address fields
- Recent searches persistence
- Additional providers (Loqate, OS Places)
- Option to pin recent addresses or show history
- Configurable minimum search length
- Localization resource files

## License

Internal project sample.

## Bundle Optimization

Originally the control imported icons from `@fluentui/react-icons`, which pulled in large aggregated icon chunks causing Babel to emit deoptimization warnings ("code generator has deoptimised the styling of..."). These icon bundles also inflated the output size (~1.5 MiB unminified).

To resolve this, the icon dependencies were replaced with two lightweight inline SVG components (`SearchIcon`, `CloseIcon`) defined directly inside `index.ts`. This eliminated the large icon chunks and reduced the final control bundle to ~25 KiB, removing the Babel deopt notes while preserving the visual design.

If additional icons are required later, prefer adding minimal inline SVGs or a micro icon subset rather than reintroducing the full Fluent UI icon package.
