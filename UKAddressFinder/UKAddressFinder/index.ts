import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import {
  createProvider,
  AddressResult,
  AddressProvider,
} from "./AddressProvider";
import {
  Input,
  Spinner,
  Card,
  CardHeader,
  Button,
  Body1,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { InputProps, InputOnChangeData } from "@fluentui/react-components";

// Enhanced Search Icon Component with better design
const SearchIcon: React.FC = () =>
  React.createElement(
    "svg",
    {
      width: 20,
      height: 20,
      viewBox: "0 0 20 20",
      role: "img",
      "aria-hidden": true,
      fill: "none",
    },
    React.createElement("path", {
      d: "m8.5 3a5.5 5.5 0 0 1 4.383 8.823l3.896 3.896a.75.75 0 0 1-.976 1.134l-.084-.073-3.896-3.896A5.5 5.5 0 1 1 8.5 3Zm0 1.5a4 4 0 1 0 0 8 4 4 0 0 0 0-8Z",
      fill: "currentColor",
    })
  );
const CloseIcon: React.FC = () =>
  React.createElement(
    "svg",
    {
      width: 16,
      height: 16,
      viewBox: "0 0 16 16",
      role: "img",
      "aria-hidden": true,
    },
    React.createElement("path", {
      fill: "currentColor",
      d: "M3.404 2.997L8 7.593l4.596-4.596.707.707L8.707 8.3l4.596 4.596-.707.707L8 9.007l-4.596 4.596-.707-.707L7.293 8.3 2.697 3.704l.707-.707z",
    })
  );
const useStyles = makeStyles({
  container: {
    width: "100%",
    position: "relative",
    fontFamily: tokens.fontFamilyBase,
  },
  dropdown: {
    listStyle: "none",
    margin: 0,
    padding: "4px 0",
    position: "absolute",
    top: "100%",
    left: 0,
    right: 0,
    zIndex: 1001,
    backgroundColor: "#ffffff",
    border: "1px solid #d1d1d1",
    borderRadius: "4px",
    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.15)",
    maxHeight: "280px",
    overflowY: "auto",
  },
  item: {
    cursor: "pointer",
    padding: "10px 12px",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: "#333333",
    "&:hover": {
      backgroundColor: "#f5f5f5",
    },
  },
  itemActive: {
    backgroundColor: "#e6f3ff",
    color: "#0078d4",
    "&:hover": {
      backgroundColor: "#e6f3ff",
    },
  },
  manual: {
    marginTop: "16px",
  },
  loadingItem: {
    padding: "10px 12px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    fontStyle: "italic",
  },
  noResultsItem: {
    padding: "10px 12px",
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    fontStyle: "italic",
    textAlign: "center",
  },
  searchInput: {
    width: "100%",
  },
  manualCard: {
    marginTop: "16px",
    padding: "20px",
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow8,
  },
  manualInputGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
    marginTop: "16px",
  },
  manualInputFull: {
    gridColumn: "1 / -1",
  },
  errorMessage: {
    marginTop: "6px",
    padding: "8px 12px",
    color: tokens.colorPaletteRedForeground1,
    backgroundColor: tokens.colorPaletteRedBackground1,
    fontSize: tokens.fontSizeBase200,
    borderRadius: tokens.borderRadiusSmall,
    border: `1px solid ${tokens.colorPaletteRedBorder1}`,
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  enhancedButton: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
});

interface AddressFinderProps {
  context: ComponentFramework.Context<IInputs>;
  onSelect: (addr: AddressResult) => void;
  provider: AddressProvider | null;
}

// Narrow subset of Input props we actually use to avoid casting to any
type MinimalInputProps = Pick<
  InputProps,
  | "placeholder"
  | "value"
  | "onChange"
  | "onFocus"
  | "onKeyDown"
  | "contentBefore"
  | "className"
> & {
  "aria-autocomplete"?: string;
  "aria-controls"?: string;
  "aria-expanded"?: boolean;
  ref?: React.RefObject<HTMLInputElement>;
};

const TypedInput = Input as unknown as React.FC<MinimalInputProps>;

const AddressFinderComponent: React.FC<AddressFinderProps> = ({
  context,
  onSelect,
  provider,
}) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = React.useState("");
  const [results, setResults] = React.useState<AddressResult[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [showDropdown, setShowDropdown] = React.useState(false);
  const [manual, setManual] = React.useState(false);
  const [highlight, setHighlight] = React.useState<number>(-1);
  const [manualAddress, setManualAddress] = React.useState<
    Partial<AddressResult>
  >({});
  const debounceRef = React.useRef<number | undefined>(undefined);
  const abortRef = React.useRef<AbortController | null>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const dropdownRef = React.useRef<HTMLUListElement>(null);
  const placeholder: string =
    (context.parameters as unknown as { placeholder?: { raw?: string } })
      .placeholder?.raw ?? "Search UK address";

  // Close dropdown when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        inputRef.current &&
        dropdownRef.current &&
        !inputRef.current.contains(event.target as Node) &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setShowDropdown(false);
      }
    };

    if (showDropdown) {
      document.addEventListener("mousedown", handleClickOutside);
      return () =>
        document.removeEventListener("mousedown", handleClickOutside);
    }
  }, [showDropdown]);

  React.useEffect(() => {
    if (!provider) return; // manual entry only
    if (searchTerm.trim().length < 3) {
      setResults([]);
      setShowDropdown(false);
      setError(null);
      return;
    }
    setLoading(true);
    setError(null);
    if (debounceRef.current) window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => {
      abortRef.current?.abort();
      const controller = new AbortController();
      abortRef.current = controller;
      void (async () => {
        try {
          const res = await provider.searchAddresses(
            searchTerm,
            controller.signal
          );
          if (!controller.signal.aborted) {
            setResults(res);
            setShowDropdown(res.length > 0);
            setError(null);
            setHighlight(-1);
          }
        } catch (e) {
          if (
            (e as Error).name !== "AbortError" &&
            !controller.signal.aborted
          ) {
            setError("Search failed. Please try again.");
            setResults([]);
            setShowDropdown(false);
          }
        } finally {
          if (!controller.signal.aborted) {
            setLoading(false);
          }
        }
      })();
    }, 300);
    return () => {
      if (debounceRef.current) window.clearTimeout(debounceRef.current);
    };
  }, [searchTerm, provider]);

  const select = async (r: AddressResult) => {
    if (provider && r.addressLine1 === "") {
      try {
        setLoading(true);
        const detail = await provider.getAddressDetails(r.id);
        onSelect(detail);
      } catch {
        onSelect(r); // fallback
      } finally {
        setLoading(false);
      }
    } else {
      onSelect(r);
    }
    setShowDropdown(false);
    setSearchTerm(r.displayText);
  };

  const handleKey: React.KeyboardEventHandler<HTMLInputElement> = (e) => {
    if (!showDropdown || results.length === 0) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setHighlight((h) => Math.min(results.length - 1, h + 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setHighlight((h) => Math.max(-1, h - 1));
    } else if (e.key === "Enter" && highlight >= 0) {
      e.preventDefault();
      void select(results[highlight]);
    } else if (e.key === "Escape") {
      e.preventDefault();
      setShowDropdown(false);
      setHighlight(-1);
    }
  };

  const manualEntry = () => {
    setManual(true);
    setShowDropdown(false);
    setSearchTerm("");
    setManualAddress({
      id: "manual",
      displayText: "",
      addressLine1: "",
      addressLine2: "",
      city: "",
      county: "",
      postcode: "",
      country: "United Kingdom",
    });
  };

  const updateManualField = (field: keyof AddressResult, value: string) => {
    const updated = { ...manualAddress, [field]: value };
    setManualAddress(updated);
    // Update display text for preview
    const parts = [
      updated.addressLine1,
      updated.addressLine2,
      updated.city,
      updated.county,
      updated.postcode,
    ].filter(Boolean);
    const displayText = parts.join(", ");
    onSelect({
      id: "manual",
      displayText,
      addressLine1: updated.addressLine1 ?? "",
      addressLine2: updated.addressLine2 ?? "",
      city: updated.city ?? "",
      county: updated.county ?? "",
      postcode: updated.postcode ?? "",
      country: updated.country ?? "United Kingdom",
    });
  };

  // helper to emphasize matched portion
  const renderSuggestion = (text: string) => {
    if (!searchTerm) return text;
    const idx = text.toLowerCase().indexOf(searchTerm.toLowerCase());
    if (idx === -1) return text;
    return React.createElement(
      React.Fragment,
      null,
      text.substring(0, idx),
      React.createElement(
        "strong",
        null,
        text.substring(idx, idx + searchTerm.length)
      ),
      text.substring(idx + searchTerm.length)
    );
  };

  return React.createElement(
    "div",
    {
      className: styles.container,
      "aria-label": "UK Address Finder",
    },
    !manual &&
      React.createElement(
        "div",
        { style: { position: "relative" } },
        React.createElement(TypedInput, {
          ref: inputRef,
          className: styles.searchInput,
          contentBefore: React.createElement(SearchIcon, null),
          placeholder,
          value: searchTerm,
          onChange: (
            _: React.ChangeEvent<HTMLInputElement>,
            d: InputOnChangeData
          ) => setSearchTerm(d.value),
          onFocus: () => {
            if (results.length > 0 && searchTerm.length >= 3) {
              setShowDropdown(true);
            }
          },
          onKeyDown: handleKey,
          "aria-autocomplete": "list",
          "aria-controls": "uk-address-results",
          "aria-expanded": showDropdown,
        }),
        showDropdown &&
          React.createElement(
            "ul",
            {
              ref: dropdownRef,
              id: "uk-address-results",
              role: "listbox",
              className: styles.dropdown,
              "aria-live": "polite",
            },
            loading &&
              React.createElement(
                "li",
                { className: styles.loadingItem, role: "status" },
                React.createElement(Spinner, { size: "extra-tiny" }),
                "Searching addresses..."
              ),
            !loading &&
              results.length === 0 &&
              searchTerm.length >= 3 &&
              React.createElement(
                "li",
                { className: styles.noResultsItem, role: "note" },
                "No addresses found. Try a different search term."
              ),
            !loading &&
              results.map((r, i) =>
                React.createElement(
                  "li",
                  {
                    key: r.id,
                    role: "option",
                    "aria-selected": highlight === i,
                    className:
                      styles.item +
                      (highlight === i ? " " + styles.itemActive : ""),
                    onMouseDown: (e: React.MouseEvent) => {
                      e.preventDefault();
                      void select(r);
                    },
                    onMouseEnter: () => setHighlight(i),
                    onMouseLeave: () => setHighlight(-1),
                  },
                  renderSuggestion(r.displayText)
                )
              )
          )
      ),
    error &&
      React.createElement(
        "div",
        { className: styles.errorMessage, role: "alert" },
        error
      ),
    !manual &&
      React.createElement(
        "div",
        { className: styles.manual },
        React.createElement(
          Button,
          {
            size: "small",
            appearance: "secondary",
            onClick: manualEntry,
            className: styles.enhancedButton,
            icon: React.createElement(
              "svg",
              {
                width: 16,
                height: 16,
                viewBox: "0 0 16 16",
                fill: "currentColor",
              },
              React.createElement("path", {
                d: "M3 4.5a.5.5 0 0 1 .5-.5h9a.5.5 0 0 1 0 1h-9a.5.5 0 0 1-.5-.5zM3 8a.5.5 0 0 1 .5-.5h9a.5.5 0 0 1 0 1h-9A.5.5 0 0 1 3 8zm0 3.5a.5.5 0 0 1 .5-.5h6a.5.5 0 0 1 0 1h-6a.5.5 0 0 1-.5-.5z",
              })
            ),
          },
          "Enter address manually"
        )
      ),
    manual &&
      React.createElement(
        Card,
        { className: styles.manualCard, appearance: "filled" },
        React.createElement(CardHeader, {
          header: React.createElement(Body1, null, "Manual Address Entry"),
          action: React.createElement(Button, {
            "aria-label": "Close manual entry",
            size: "small",
            appearance: "subtle",
            icon: React.createElement(CloseIcon, null),
            onClick: () => {
              setManual(false);
              setManualAddress({});
            },
          }),
        }),
        React.createElement(
          "div",
          { className: styles.manualInputGrid },
          React.createElement(TypedInput, {
            className: styles.manualInputFull,
            placeholder: "Address Line 1*",
            value: manualAddress.addressLine1 ?? "",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("addressLine1", d.value),
          }),
          React.createElement(TypedInput, {
            className: styles.manualInputFull,
            placeholder: "Address Line 2",
            value: manualAddress.addressLine2 ?? "",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("addressLine2", d.value),
          }),
          React.createElement(TypedInput, {
            placeholder: "City/Town",
            value: manualAddress.city ?? "",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("city", d.value),
          }),
          React.createElement(TypedInput, {
            placeholder: "County",
            value: manualAddress.county ?? "",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("county", d.value),
          }),
          React.createElement(TypedInput, {
            placeholder: "Postcode",
            value: manualAddress.postcode ?? "",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("postcode", d.value),
          }),
          React.createElement(TypedInput, {
            placeholder: "Country",
            value: manualAddress.country ?? "United Kingdom",
            onChange: (
              _: React.ChangeEvent<HTMLInputElement>,
              d: InputOnChangeData
            ) => updateManualField("country", d.value),
          })
        )
      )
  );
};

export class UKAddressFinder
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private selectedAddress: AddressResult | null = null;
  private provider: AddressProvider | null = null;

  /**
   * Empty constructor.
   */
  constructor() {
    // Empty
  }

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.provider = createProvider(context);
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    // provider might change if apiKey updated
    this.provider = createProvider(context);
    return React.createElement(AddressFinderComponent, {
      context,
      provider: this.provider,
      onSelect: (addr: AddressResult) => {
        this.selectedAddress = addr;
        this.notifyOutputChanged();
      },
    });
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    const outputs: IOutputs = {} as IOutputs;
    if (this.selectedAddress) {
      /* eslint-disable @typescript-eslint/no-explicit-any,@typescript-eslint/no-unsafe-member-access */
      (outputs as any).addressLine1 = this.selectedAddress.addressLine1;
      (outputs as any).addressLine2 = this.selectedAddress.addressLine2;
      (outputs as any).city = this.selectedAddress.city;
      (outputs as any).county = this.selectedAddress.county;
      (outputs as any).postcode = this.selectedAddress.postcode;
      (outputs as any).country = this.selectedAddress.country;
      /* eslint-enable @typescript-eslint/no-explicit-any,@typescript-eslint/no-unsafe-member-access */
    }
    return outputs;
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}
