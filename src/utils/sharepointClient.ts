import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import type {
  DiscoveredList,
  SharePointClient,
} from "../types";

/** Wraps fetch with an AbortController timeout (default 30s) */
async function fetchWithTimeout(url: string | URL | Request, options: RequestInit = {}, timeoutMs = 30000): Promise<Response> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await globalThis.fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeoutId);
  }
}

// Module-level digest cache
let _digestCache: string | null = null;
let _digestExpiry: number = 0;

function spScope(): string {
  const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");
  try {
    const origin = new URL(SP_SITE_URL).origin;
    return `${origin}/AllSites.Manage`;
  } catch {
    return "https://graph.microsoft.com/.default";
  }
}

async function getToken(
  instance: IPublicClientApplication,
  accounts: AccountInfo[]
): Promise<string> {
  if (accounts.length === 0) {
    throw new Error("No accounts found. User must sign in first.");
  }

  try {
    const response = await instance.acquireTokenSilent({
      scopes: [spScope()],
      account: accounts[0],
    });
    return response.accessToken;
  } catch {
    const response = await instance.acquireTokenPopup({
      scopes: [spScope()],
    });
    return response.accessToken;
  }
}

async function getDigest(token: string): Promise<string> {
  const now = Date.now();
  if (_digestCache && now < _digestExpiry) {
    return _digestCache;
  }

  const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");

  const response = await fetchWithTimeout(`${SP_SITE_URL}/_api/contextinfo`, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to get digest: ${response.status}`);
  }

  const data = await response.json();
  const formDigestValue = data?.FormDigestValue;

  if (typeof formDigestValue !== "string") {
    throw new Error("Invalid digest response");
  }

  _digestCache = formDigestValue;
  _digestExpiry = now + 1800000; // 30 minutes

  return formDigestValue;
}

function buildSelect(columns: string[]): string {
  const cols = new Set(["Id", ...columns]);
  return [...cols].join(",");
}

export function createSpClient(
  instance: IPublicClientApplication,
  accounts: AccountInfo[]
): SharePointClient {
  const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");

  async function acquireToken(): Promise<string> {
    return getToken(instance, accounts);
  }

  function getCurrentUserEmail(): string {
    if (accounts.length === 0) {
      throw new Error("No accounts found");
    }
    const email = accounts[0].username;
    if (!email) {
      throw new Error("No email found in account");
    }
    return email;
  }

  async function discoverLists(): Promise<DiscoveredList[]> {
    const token = await acquireToken();
    const response = await fetchWithTimeout(
      `${SP_SITE_URL}/_api/web/lists?$select=Title,Id,ItemCount,Created,Hidden,BaseTemplate,BaseType,IsCatalog,IsSiteAssetsLibrary,IsApplicationList,IsSystemList,NoCrawl`,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          Authorization: `Bearer ${token}`,
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to discover lists: ${response.status}`);
    }

    const data = await response.json();
    const results = data.value;

    if (!Array.isArray(results)) {
      return [];
    }

    return results.map((list: Record<string, unknown>) => ({
      title: String(list.Title || ""),
      id: String(list.Id || ""),
      itemCount: Number(list.ItemCount) || 0,
      created: String(list.Created || ""),
      hidden: Boolean(list.Hidden),
      baseTemplate: Number(list.BaseTemplate) || 0,
      baseType: Number(list.BaseType) || 0,
      isCatalog: Boolean(list.IsCatalog),
      isSiteAssetsLibrary: Boolean(list.IsSiteAssetsLibrary),
      isApplicationList: Boolean(list.IsApplicationList),
      isSystemList: Boolean(list.IsSystemList),
      noCrawl: Boolean(list.NoCrawl),
    }));
  }

  async function queryList(
    listName: string,
    options?: Record<string, unknown>
  ): Promise<Record<string, unknown>[]> {
    const token = await acquireToken();
    const params = new URLSearchParams();
    const selectInput = options?.select;
    const selectArr = Array.isArray(selectInput) ? selectInput : typeof selectInput === "string" ? [selectInput] : [];
    const selectCols = buildSelect([...selectArr, "Author/Id,Author/Email"]);
    params.set("$select", selectCols);
    if (options?.filter) params.set("$filter", options.filter as string);
    if (options?.orderby) params.set("$orderby", options.orderby as string);
    params.set("$top", String(options?.top ?? 500));
    params.set("$expand", "Author/Id,Author/Email");

    const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?${params}`;
    const response = await fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to query list: ${response.status}`);
    }

    const data = await response.json();
    const items = data.value || [];

    if (!Array.isArray(items) || items.length === 0) {
      return [];
    }

    return items.map((item: Record<string, unknown>) => ({
      ...item,
      _authorEmail: (item.Author as Record<string, unknown>)?.Email as string || "",
    }));
  }

  async function queryListByGuid(
    listGuid: string,
    options?: Record<string, unknown>
  ): Promise<Record<string, unknown>[]> {
    const token = await acquireToken();
    const params = new URLSearchParams();
    const selectInput = options?.select;
    const selectArr = Array.isArray(selectInput) ? selectInput : typeof selectInput === "string" ? [selectInput] : [];
    const selectCols = buildSelect([...selectArr, "Author/Id,Author/Email"]);
    params.set("$select", selectCols);
    if (options?.filter) params.set("$filter", options.filter as string);
    if (options?.orderby) params.set("$orderby", options.orderby as string);
    params.set("$top", String(options?.top ?? 500));
    params.set("$expand", "Author/Id,Author/Email");

    const url = `${SP_SITE_URL}/_api/web/lists(guid'${listGuid}')/items?${params}`;
    const response = await fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to query list by GUID: ${response.status}`);
    }

    const data = await response.json();
    const items = data.value || [];

    if (!Array.isArray(items) || items.length === 0) {
      return [];
    }

    return items.map((item: Record<string, unknown>) => ({
      ...item,
      _authorEmail: (item.Author as Record<string, unknown>)?.Email as string || "",
    }));
  }

  async function queryListByEmail(
    listName: string,
    userEmail: string,
    options?: Record<string, unknown>
  ): Promise<Record<string, unknown>[]> {
    const token = await acquireToken();
    const params = new URLSearchParams();
    const selectCols = buildSelect([...(options?.select as string[] || []), "Author/Id,Author/Email", "FormID", "NumberOfApprovalLayers", "FormStatus"]);
    params.set("$select", selectCols);
    if (options?.filter) params.set("$filter", options.filter as string);
    if (options?.orderby) params.set("$orderby", options.orderby as string);
    params.set("$top", String(options?.top ?? 500));
    params.set("$expand", "Author/Id,Author/Email");

    const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?${params}`;
    const response = await fetchWithTimeout(url, {
      headers: {
        Accept: "application/json;odata=nometadata",
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    const items = data.value || [];

    if (!Array.isArray(items) || items.length === 0) {
      return [];
    }

    const userEmailLower = userEmail.toLowerCase();
    return items
      .filter((item: Record<string, unknown>) => {
        const authorEmail = ((item.Author as Record<string, unknown>)?.Email as string) || "";
        return authorEmail.toLowerCase() === userEmailLower;
      })
      .map((item: Record<string, unknown>) => ({
        ...item,
        _authorEmail: (item.Author as Record<string, unknown>)?.Email as string || "",
      }));
  }

  async function isGroupMember(groupName: string): Promise<boolean> {
    const token = await acquireToken();
    const email = getCurrentUserEmail();

    try {
      const response = await fetchWithTimeout(
        `${SP_SITE_URL}/_api/web/sitegroups/getByName('${encodeURIComponent(groupName)}')/users?$select=LoginName,Email`,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            Authorization: `Bearer ${token}`,
          },
        }
      );

      if (!response.ok) return false;

      const data = await response.json();
      const members: Record<string, unknown>[] = data.value || [];
      const userEmail = email.toLowerCase();

      return members.some(
        (m) =>
          (String(m.Email || "")).toLowerCase() === userEmail ||
          (String(m.LoginName || "")).toLowerCase().split("|").pop() === userEmail
      );
    } catch {
      return false;
    }
  }

  async function listExists(title: string): Promise<boolean> {
    const token = await acquireToken();

    try {
      const response = await fetchWithTimeout(
        `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Id`,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            Authorization: `Bearer ${token}`,
          },
        }
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  async function createList(
    title: string,
    description?: string
  ): Promise<string | null> {
    const token = await acquireToken();
    const digest = await getDigest(token);

    const body = {
      __metadata: { type: "SP.List" },
      Title: title,
      Description: description || "",
      BaseTemplate: 100,
    };

    const response = await fetchWithTimeout(`${SP_SITE_URL}/_api/web/lists`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      return null;
    }

    const data = await response.json();
    return data?.Id || data?.id || null;
  }

  async function addColumn(
    listTitle: string,
    internalName: string,
    fieldTypeKind: number,
    isMultiLine?: boolean
  ): Promise<void> {
    const token = await acquireToken();
    const digest = await getDigest(token);

    const body: Record<string, unknown> = {
      __metadata: { type: "SP.Field" },
      Title: internalName,
      FieldTypeKind: fieldTypeKind,
    };

    if (fieldTypeKind === 3 && isMultiLine) {
      body["RichText"] = false;
      body["NumLines"] = 6;
    }

    const response = await fetchWithTimeout(
      `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify(body),
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to add column: ${response.status}`);
    }
  }

  async function upsertListItem(
    listTitle: string,
    filterExpr: string,
    body: Record<string, unknown>
  ): Promise<{ updated: boolean; id: string }> {
    const token = await acquireToken();

    // Try to find existing item
    const existing = await queryList(listTitle, { filter: filterExpr, top: 1 });

    if (existing.length > 0) {
      // Update existing
      const item = existing[0];
      const itemId = String(item.Id || "");
      const digest = await getDigest(token);

      const updateBody = {
        __metadata: { type: "SP.Data." + listTitle.replace(/\s/g, "_x0020_") + "ListItem" },
        ...body,
      };

      const response = await fetchWithTimeout(
        `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=nometadata",
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest,
            "X-HTTP-Method": "MERGE",
            "If-Match": "*",
          },
          body: JSON.stringify(updateBody),
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to update item: ${response.status}`);
      }

      return { updated: true, id: itemId };
    } else {
      // Create new
      const digest = await getDigest(token);

      const createBody = {
        __metadata: { type: "SP.Data." + listTitle.replace(/\s/g, "_x0020_") + "ListItem" },
        ...body,
      };

      const response = await fetchWithTimeout(
        `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=nometadata",
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest,
          },
          body: JSON.stringify(createBody),
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to create item: ${response.status}`);
      }

      const data = await response.json();
      return {
        updated: false,
        id: data?.Id || "",
      };
    }
  }

  async function deleteListItemsWhere(
    listTitle: string,
    filterExpr: string
  ): Promise<number> {
    const token = await acquireToken();
    const items = await queryList(listTitle, { filter: filterExpr, top: 500 });
    let deleted = 0;

    for (const item of items) {
      const itemId = String(item.Id || "");
      if (!itemId) continue;

      try {
        const digest = await getDigest(token);
        const response = await fetchWithTimeout(
      `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=nometadata",
              Authorization: `Bearer ${token}`,
              "X-RequestDigest": digest,
              "X-HTTP-Method": "DELETE",
              "If-Match": "*",
            },
          }
        );

        if (response.ok) {
          deleted++;
        }
      } catch {
        // Continue with other deletions
      }
    }

    return deleted;
  }

  async function getSiteUsers(): Promise<{ email: string; name: string }[]> {
    const token = await acquireToken();

    try {
      const response = await fetchWithTimeout(
        `${SP_SITE_URL}/_api/web/siteusers?$select=Email,Title&$filter=PrincipalType eq 1`,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            Authorization: `Bearer ${token}`,
          },
        }
      );

      if (!response.ok) {
        return [];
      }

      const data = await response.json();
      return (data.value || [])
        .filter((u: Record<string, unknown>) => u.Email)
        .map((u: Record<string, unknown>) => ({
          email: String(u.Email || ""),
          name: String(u.Title || ""),
        }));
    } catch {
      return [];
    }
  }

  return {
    discoverLists,
    queryList,
    queryListByGuid,
    queryListByEmail,
    isGroupMember,
    getCurrentUserEmail,
    acquireToken,
    listExists,
    createList,
    addColumn,
    upsertListItem,
    deleteListItemsWhere,
    getSiteUsers,
  };
}
