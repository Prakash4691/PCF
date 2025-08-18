import { IInputs } from "./generated/ManifestTypes";

export interface AddressResult {
  id: string;
  displayText: string;
  addressLine1: string;
  addressLine2?: string;
  city: string;
  county: string;
  postcode: string;
  country: string;
}

export interface AddressProvider {
  searchAddresses(
    query: string,
    signal?: AbortSignal
  ): Promise<AddressResult[]>;
  getAddressDetails(id: string, signal?: AbortSignal): Promise<AddressResult>;
}

interface CacheEntry<T> {
  value: T;
  timestamp: number;
}

export class GetAddressProvider implements AddressProvider {
  private apiKey: string;
  private baseUrl = "https://api.getAddress.io";
  private searchCache: Map<string, CacheEntry<AddressResult[]>> = new Map<
    string,
    CacheEntry<AddressResult[]>
  >();
  private detailCache: Map<string, CacheEntry<AddressResult>> = new Map<
    string,
    CacheEntry<AddressResult>
  >();
  private cacheTtlMs = 5 * 60 * 1000; // 5 minutes

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  private isFresh<T>(entry?: CacheEntry<T>): entry is CacheEntry<T> {
    return !!entry && Date.now() - entry.timestamp < this.cacheTtlMs;
  }

  async searchAddresses(
    query: string,
    signal?: AbortSignal
  ): Promise<AddressResult[]> {
    const q = query.trim();
    if (!q) return [];
    // Accept postcode or address line search with 2+ chars if pattern matches
    const postcodePattern = /^[A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}$/i;
    const allowShort =
      postcodePattern.test(q.replace(/\s+/g, "")) || q.length >= 3;
    if (!allowShort) return [];
    const cached = this.searchCache.get(q.toLowerCase());
    if (this.isFresh(cached)) return cached.value;

    const controller = signal ? undefined : new AbortController();
    const finalSignal = signal ?? controller?.signal;

    interface AutocompleteSuggestion {
      id: string;
      address: string;
    }
    interface AutocompleteResponse {
      suggestions?: AutocompleteSuggestion[];
    }
    // Add 'all=true' for broader results, and top=20 for more suggestions
    const url = `${this.baseUrl}/autocomplete/${encodeURIComponent(
      q
    )}?api-key=${encodeURIComponent(this.apiKey)}&all=true&top=20`;
    const resp = await this.fetchWithTimeout(
      url,
      { signal: finalSignal },
      5000
    );
    if (!resp.ok) throw new Error(`Search failed: ${resp.status}`);
    const data: AutocompleteResponse =
      (await resp.json()) as AutocompleteResponse;
    const suggestions: AutocompleteSuggestion[] = data.suggestions ?? [];
    const results: AddressResult[] = suggestions.map((s) => ({
      id: s.id,
      displayText: s.address,
      addressLine1: "",
      addressLine2: "",
      city: "",
      county: "",
      postcode: "",
      country: "United Kingdom",
    }));
    this.searchCache.set(q.toLowerCase(), {
      value: results,
      timestamp: Date.now(),
    });
    return results;
  }

  async getAddressDetails(
    id: string,
    signal?: AbortSignal
  ): Promise<AddressResult> {
    const cached = this.detailCache.get(id);
    if (this.isFresh(cached)) return Promise.resolve(cached.value);

    const controller = signal ? undefined : new AbortController();
    const finalSignal = signal ?? controller?.signal;

    const url = `${this.baseUrl}/get/${encodeURIComponent(
      id
    )}?api-key=${encodeURIComponent(this.apiKey)}`;
    const resp = await this.fetchWithTimeout(
      url,
      { signal: finalSignal },
      5000
    );
    if (!resp.ok) throw new Error(`Detail failed: ${resp.status}`);
    // Define a loose type but not any
    interface DetailResponse {
      line_1?: string;
      line_2?: string;
      thoroughfare?: string;
      building_and_street?: string[];
      sub_building_name?: string;
      town_or_city?: string;
      post_town?: string;
      county?: string;
      county_name?: string;
      postcode?: string;
      post_code?: string;
    }
    const data: DetailResponse = (await resp.json()) as DetailResponse;

    const line1 =
      data.line_1 ?? data.thoroughfare ?? data.building_and_street?.[0] ?? "";
    const line2 =
      data.line_2 ??
      data.building_and_street?.[1] ??
      data.sub_building_name ??
      "";
    const city = data.town_or_city ?? data.post_town ?? "";
    const county = data.county ?? data.county_name ?? "";
    const postcode = data.postcode ?? data.post_code ?? "";

    // Build displayText for dropdown and output fields
    const displayParts = [line1, line2, city, county, postcode].filter(Boolean);
    const result: AddressResult = {
      id,
      displayText: displayParts.join(", "),
      addressLine1: line1,
      addressLine2: line2,
      city,
      county,
      postcode,
      country: "United Kingdom",
    };
    this.detailCache.set(id, { value: result, timestamp: Date.now() });
    return result;
  }

  private async fetchWithTimeout(
    input: RequestInfo,
    init: RequestInit & { timeoutMs?: number } = {},
    timeoutMs = 5000
  ): Promise<Response> {
    const { timeoutMs: override, signal, ...rest } = init;
    const to = override ?? timeoutMs;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), to);
    try {
      const combinedSignal = signal
        ? this.mergeSignals(signal, controller.signal)
        : controller.signal;
      return await fetch(input, { ...rest, signal: combinedSignal });
    } finally {
      clearTimeout(timer);
    }
  }

  private mergeSignals(
    signalA: AbortSignal,
    signalB: AbortSignal
  ): AbortSignal {
    if (signalA.aborted) return signalA;
    if (signalB.aborted) return signalB;
    const controller = new AbortController();
    const abort = () => controller.abort();
    signalA.addEventListener("abort", abort);
    signalB.addEventListener("abort", abort);
    return controller.signal;
  }
}

// Azure Maps provider using Fuzzy Search API
// Docs: https://learn.microsoft.com/rest/api/maps/search/get-search-fuzzy
// We reuse the apiKey property (x-ms-client-id) as per user instruction.
export class AzureMapsProvider implements AddressProvider {
  private apiKey: string; // used as subscription-key OR x-ms-client-id depending on Azure setup; here we'll send as subscription-key header param
  private baseUrl = "https://atlas.microsoft.com";
  private searchCache: Map<string, CacheEntry<AddressResult[]>> = new Map<
    string,
    CacheEntry<AddressResult[]>
  >();
  private detailCache: Map<string, CacheEntry<AddressResult>> = new Map<
    string,
    CacheEntry<AddressResult>
  >();
  private cacheTtlMs = 5 * 60 * 1000; // 5 minutes

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  private isFresh<T>(entry?: CacheEntry<T>): entry is CacheEntry<T> {
    return !!entry && Date.now() - entry.timestamp < this.cacheTtlMs;
  }

  async searchAddresses(
    query: string,
    signal?: AbortSignal
  ): Promise<AddressResult[]> {
    const q = query.trim();
    if (!q) return [];
    if (q.length < 3) return [];
    const cached = this.searchCache.get(q.toLowerCase());
    if (this.isFresh(cached)) return cached.value;

    const controller = signal ? undefined : new AbortController();
    const finalSignal = signal ?? controller?.signal;

    // Azure Maps Fuzzy Search endpoint: /search/fuzzy
    // Required query params: 'api-version', 'query'
    // We'll restrict countrySet=GB to focus on UK results and limit size=20 for parity.
    const params = new URLSearchParams({
      "api-version": "1.0",
      query: q,
      countrySet: "GB",
      limit: "20",
    });
    const url = `${this.baseUrl}/search/fuzzy/json?${params.toString()}`;
    const resp = await this.fetchWithTimeout(url, {
      signal: finalSignal,
      headers: {
        // Subscription key header (if using key). Some setups use 'subscription-key' or 'x-ms-client-id'. We'll send both for flexibility.
        "subscription-key": this.apiKey,
        "x-ms-client-id": this.apiKey,
      },
    });
    if (!resp.ok) throw new Error(`Search failed: ${resp.status}`);
    interface FuzzyResponse {
      results?: {
        id?: string;
        address?: {
          freeformAddress?: string;
          streetNumber?: string;
          streetName?: string;
          municipality?: string; // town/city
          countrySubdivision?: string; // county
          postalCode?: string;
          country?: string;
        };
      }[];
    }
    const data: FuzzyResponse = (await resp.json()) as FuzzyResponse;
    const results: AddressResult[] = (data.results ?? []).map((r, idx) => {
      const addr = r.address ?? {};
      // We don't have separate detail endpoint for now; we'll treat fuzzy search as final output.
      const line1 =
        addr.freeformAddress?.split(",")[0]?.trim() ??
        addr.streetName ??
        addr.freeformAddress ??
        "";
      const line2 = ""; // Not reliably provided; could attempt to parse leftover parts but keep simple.
      const city = addr.municipality ?? "";
      const county = addr.countrySubdivision ?? "";
      const postcode = addr.postalCode ?? "";
      const country = addr.country ?? "United Kingdom";
      const displayText =
        addr.freeformAddress ??
        [line1, city, county, postcode].filter(Boolean).join(", ");
      return {
        id: r.id ?? `az-${idx}-${displayText}`,
        displayText,
        addressLine1: line1,
        addressLine2: line2,
        city,
        county,
        postcode,
        country,
      } as AddressResult;
    });
    this.searchCache.set(q.toLowerCase(), {
      value: results,
      timestamp: Date.now(),
    });
    return results;
  }

  getAddressDetails(id: string, _signal?: AbortSignal): Promise<AddressResult> {
    // For Azure Maps we already have full details from search; use cache.
    const cached = this.detailCache.get(id);
    if (this.isFresh(cached)) return Promise.resolve(cached.value);
    // Attempt to find it from search caches
    for (const entry of this.searchCache.values()) {
      const found = entry.value.find((r) => r.id === id);
      if (found) {
        this.detailCache.set(id, { value: found, timestamp: Date.now() });
        return Promise.resolve(found);
      }
    }
    return Promise.reject(new Error("Address not found in cache"));
  }

  private async fetchWithTimeout(
    input: RequestInfo,
    init: RequestInit & { timeoutMs?: number } = {},
    timeoutMs = 5000
  ): Promise<Response> {
    const { timeoutMs: override, signal, ...rest } = init;
    const to = override ?? timeoutMs;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), to);
    try {
      const combinedSignal = signal
        ? this.mergeSignals(signal, controller.signal)
        : controller.signal;
      return await fetch(input, { ...rest, signal: combinedSignal });
    } finally {
      clearTimeout(timer);
    }
  }

  private mergeSignals(
    signalA: AbortSignal,
    signalB: AbortSignal
  ): AbortSignal {
    if (signalA.aborted) return signalA;
    if (signalB.aborted) return signalB;
    const controller = new AbortController();
    const abort = () => controller.abort();
    signalA.addEventListener("abort", abort);
    signalB.addEventListener("abort", abort);
    return controller.signal;
  }
}

export function createProvider(
  context: ComponentFramework.Context<IInputs>
): AddressProvider | null {
  /* eslint-disable @typescript-eslint/no-unsafe-member-access,@typescript-eslint/no-explicit-any */
  const anyParams: Record<string, unknown> =
    context.parameters as unknown as Record<string, unknown>;
  const apiKey = ((anyParams as any).apiKey?.raw as string | undefined) ?? "";
  const provider =
    ((anyParams as any).apiProvider?.raw as string | undefined) ?? "GetAddress";
  /* eslint-enable @typescript-eslint/no-unsafe-member-access,@typescript-eslint/no-explicit-any */
  if (apiKey) {
    if (provider === "GetAddress") return new GetAddressProvider(apiKey);
    if (provider === "AzureMaps") return new AzureMapsProvider(apiKey);
  }
  return null; // Manual entry fallback
}
