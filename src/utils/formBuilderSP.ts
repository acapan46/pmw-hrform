import type { FormConfig, FormLogEntry, Submission, SurveyJson } from '../types/index.ts';
import { flattenQuestions, getSpColumnKind } from './FormBuilderEngine.ts';

const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL as string || '').replace(/\/$/, '');

const DIGEST_EXPIRY_MS = 30 * 60 * 1000; // 30 minutes
let cachedDigest: string | null = null;
let digestExpiry: number | null = null;

async function getDigest(token: string): Promise<string> {
  const now = Date.now();
  if (cachedDigest && digestExpiry && now < digestExpiry) {
    return cachedDigest;
  }

  const url = `${SP_SITE_URL}/_api/contextinfo`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch request digest: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  if (!data.FormDigestValue) {
    throw new Error('No FormDigestValue returned from contextinfo endpoint');
  }

  const digestValue: string = data.FormDigestValue;
  cachedDigest = digestValue;
  digestExpiry = now + DIGEST_EXPIRY_MS;
  return digestValue;
}

export async function getFormConfig(token: string, listTitle: string): Promise<FormConfigData | null> {
  if (!await listExists(token, 'Master Form')) return null;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Title eq '${encodeURIComponent(listTitle)}'&$select=Id,Title,FormID,NumberOfApprovalLayer,Slug,CurrentVersion,IsPublished,IsPublic,ConditionField,ApprovalRules&$top=1`) as { value?: FormConfigData[] };
  return data.value?.[0] || null;
}

export async function saveFormConfig(
  config: Omit<FormConfig, 'Id' | 'Created' | 'Modified'>,
  token: string
): Promise<FormConfig> {
  const digest = await getDigest(token);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
      'X-RequestDigest': digest,
    },
    body: JSON.stringify(config),
  });

  if (!response.ok) {
    throw new Error(`Failed to save form config: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

export async function saveFormVersion(
  token: string,
  params: {
    listTitle: string;
    slug: string;
    version: string;
    surveyJson: unknown;
    meta: unknown;
    changedBy: string;
  }
): Promise<void> {
  await ensureListExists(token, 'Web Form Versions');
  const jsonStr = JSON.stringify({ surveyJson: params.surveyJson, meta: params.meta, version: params.version, savedAt: new Date().toISOString(), changedBy: params.changedBy }, null, 2);
  const existing = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(params.listTitle)}' and FormVersion eq '${encodeURIComponent(params.version)}'&$select=Id&$top=1`).catch(() => ({ value: [] })) as { value?: { Id: number }[] };
  const body = {
    Title: `${params.listTitle} v${params.version}`,
    FormTitle: params.listTitle,
    FormSlug: params.slug,
    FormVersion: params.version,
    SurveyJSON: jsonStr,
    PublishedBy: params.changedBy,
    PublishedAt: new Date().toISOString(),
  };
  if (existing.value?.length && existing.value[0].Id) {
    await spPatch(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items(${existing.value[0].Id})`, body);
  } else {
    await spPost(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items`, body);
  }
}

export async function logFormAction(
  logEntry: Omit<FormLogEntry, 'Id' | 'Timestamp'>,
  token: string
): Promise<FormLogEntry> {
  const digest = await getDigest(token);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('Form%20Builder%20Log')/items`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
      'X-RequestDigest': digest,
    },
    body: JSON.stringify(logEntry),
  });

  if (!response.ok) {
    throw new Error(`Failed to log form action: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

export async function getFormSubmissions(formId: string, token: string): Promise<Submission[]> {
  const encodedFormId = formId.replace(/'/g, "''");
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('Submissions')/items?$filter=FormId eq '${encodedFormId}'&$orderby=Created desc`;
  const response = await fetch(url, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch form submissions: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function submitFormResponse(
  formId: string,
  responseData: unknown,
  token: string
): Promise<Submission> {
  const digest = await getDigest(token);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('Submissions')/items`;
  const body = {
    FormId: formId,
    Response: JSON.stringify(responseData),
    Submitted: new Date().toISOString(),
  };
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
      'X-RequestDigest': digest,
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    throw new Error(`Failed to submit form response: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

export async function getSharePointChoices(
  listTitle: string,
  fieldName: string,
  token: string
): Promise<string[]> {
  const encodedListTitle = encodeURIComponent(listTitle);
  const encodedFieldName = encodeURIComponent(fieldName);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodedListTitle}')/fields?$filter=Title eq '${encodedFieldName}'`;
  const response = await fetch(url, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch SharePoint choices: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  const field = data.value?.[0];
  if (!field) {
    return [];
  }
  const choices = field.Choices;
  if (!choices) {
    return [];
  }
  return Array.isArray(choices) ? choices : (choices.results || []);
}

export async function getSharePointLists(token: string): Promise<{ title: string; id: string }[]> {
  const url = `${SP_SITE_URL}/_api/web/lists?$select=Id,Title,Hidden&$filter=Hidden eq false&$top=500`;
  const response = await fetch(url, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch SharePoint lists: ${response.status} ${response.statusText}`);
  }
  const data = await response.json();
  return (data.value || [])
    .filter((list: { Title?: string }) => !!list.Title)
    .map((list: { Title: string; Id: string }) => ({ title: list.Title, id: list.Id }));
}

export async function getChoiceColumnsForList(listTitle: string, token: string): Promise<{ title: string; typeKind: number; choices: string[] }[]> {
  const encodedListTitle = encodeURIComponent(listTitle);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodedListTitle}')/fields?$select=Title,FieldTypeKind,Choices&$filter=FieldTypeKind eq 6 or FieldTypeKind eq 15`;
  const response = await fetch(url, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch choice columns: ${response.status} ${response.statusText}`);
  }
  const data = await response.json();
  return (data.value || [])
    .filter((field: { Title?: string }) => !!field.Title)
    .map((field: { Title: string; FieldTypeKind: number; Choices?: { results?: string[] } | string[] }) => {
      const rawChoices = field.Choices;
      const choiceArr = Array.isArray(rawChoices)
        ? rawChoices
        : (rawChoices?.results || []);
      return { title: field.Title, typeKind: field.FieldTypeKind, choices: choiceArr };
    });
}

export function slugify(str: string): string {
  return str
    .toLowerCase()
    .trim()
    .replace(/[^a-zA-Z0-9_\s-]/g, '')
    .replace(/[\s]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

export async function checkSlugConflict(
  token: string,
  slug: string,
  excludeFormTitle?: string | null
): Promise<string | null> {
  const slugToCheck = slugify(slug);
  if (slugToCheck.length === 0) return null;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Slug eq '${encodeURIComponent(slugToCheck)}'&$select=Title,Slug&$top=5`).catch(() => ({ value: [] })) as { value?: { Title: string }[] };
  const others = (data.value || []).filter(r => r.Title !== excludeFormTitle);
  return others.length > 0 ? others[0].Title : null;
}

export async function getAllSlugs(token: string): Promise<{ Title: string; Slug: string; CurrentVersion: string }[]> {
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$select=Title,Slug,CurrentVersion&$top=500`).catch(() => ({ value: [] })) as { value?: { Title: string; Slug: string; CurrentVersion: string }[] };
  return data.value || [];
}

export async function spUploadFile(token: string, lib: string, filename: string, content: string | Uint8Array): Promise<unknown> {
  const digest = await getDigest(token);
  const r = await fetch(
    `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')/rootfolder/files/add(url='${encodeURIComponent(filename)}',overwrite=true)`,
    { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'X-RequestDigest': digest, 'Content-Type': 'application/octet-stream' }, body: (typeof content === 'string' ? new TextEncoder().encode(content) : content) as BodyInit }
  );
  if (!r.ok) { const t = await r.text(); throw new Error(`Upload ${r.status}: ${t}`); }
  return r.json().catch(() => ({}));
}

export async function getFormLog(token: string, listTitle: string): Promise<FormLogEntry[]> {
  if (!await listExists(token, 'Form Builder Log')) return [];
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Form%20Builder%20Log')/items?$filter=FormTitle eq '${encodeURIComponent(listTitle)}'&$select=EventType,ChangedBy,EventSummary,BeforeJSON,AfterJSON,EventAt,Title&$orderby=EventAt desc&$top=200`).catch(() => ({ value: [] })) as { value?: FormLogEntry[] };
  return data.value || [];
}

export async function getFormVersion(token: string, listTitle: string, version: string): Promise<{ surveyJson: unknown; meta: unknown } | null> {
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(listTitle)}' and FormVersion eq '${encodeURIComponent(version)}'&$select=SurveyJSON,FormVersion,PublishedAt,PublishedBy&$top=1`).catch(() => ({ value: [] })) as { value?: { SurveyJSON?: string }[] };
  const row = data.value?.[0];
  if (!row?.SurveyJSON) return null;
  try {
    return JSON.parse(row.SurveyJSON);
  } catch {
    return null;
  }
}

/**
 * addColumn — idempotent.
 * kind: 2=Text 3=Note 4=DateTime 6=Choice 8=Boolean 9=Number 15=MultiChoice
 * multiLine=true → SP.FieldMultiLineText (kind must be 3)
 * richText=true → Enhanced Rich Text (multiLine must be true)
 * choices → required for kind 6 (Choice) and 15 (MultiChoice)
 */
export async function addColumn(
  token: string,
  listTitle: string,
  fieldName: string,
  kind: number,
  multiLine = false,
  richText = false,
  choices?: string[]
): Promise<void> {
  // Check if already exists
  try {
    await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/getbyinternalnameortitle('${encodeURIComponent(fieldName)}')?$select=InternalName`);
    return; // already exists
  } catch {
    // Continue to create
  }

  const digest = await getDigest(token);
  const typeMap: Record<number, string> = {
    2: 'SP.Field',
    3: 'SP.FieldMultiLineText',
    4: 'SP.FieldDateTime',
    6: 'SP.FieldChoice',
    8: 'SP.Field',
    9: 'SP.FieldNumber',
    11: 'SP.FieldUrl',
    15: 'SP.FieldMultiChoice',
  };
  const body: Record<string, unknown> = {
    __metadata: { type: typeMap[kind] ?? 'SP.Field' },
    FieldTypeKind: kind,
    Title: fieldName,
    StaticName: fieldName,
  };
  if (kind === 3 || multiLine) {
    body.NumberOfLines = 6;
    body.RichText = !!richText;
  }
  if (kind === 11) {
    body.DisplayFormat = 2; // 2 = Image (0=Hyperlink, 1=Text)
  }
  if ((kind === 6 || kind === 15) && choices && choices.length > 0) {
    body.Choices = { results: choices };
  }

  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest,
    },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const text = await response.text();
    if (text.toLowerCase().includes('duplicate') || text.toLowerCase().includes('already exists')) return;
    throw new Error(`addColumn "${fieldName}" ${response.status}: ${text}`);
  }
}

export async function deleteListColumnsWhere(
  listTitle: string,
  filterExpr: string,
  token: string
): Promise<number> {
  const encodedListTitle = encodeURIComponent(listTitle);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodedListTitle}')/Fields?$filter=${encodeURIComponent(filterExpr)}`;
  const response = await fetch(url, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Authorization': `Bearer ${token}`,
    },
  });
  if (!response.ok) {
    return 0;
  }
  const data = await response.json();
  const columns = data.value || [];
  let deleted = 0;
  for (const item of columns) {
    if (!item.Id) continue;
    const encodedId = encodeURIComponent(item.Id.toString());
    const deleteUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodedListTitle}')/Fields('${encodedId}')`;
    const digest = await getDigest(token);
    const deleteResponse = await fetch(deleteUrl, {
      method: 'DELETE',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Authorization': `Bearer ${token}`,
        'X-RequestDigest': digest,
      },
    });
    if (deleteResponse.ok) {
      deleted += 1;
    }
  }
  return deleted;
}

export async function createSpList(
  token: string,
  listTitle: string,
  baseTemplate = 100,
  description = ""
): Promise<unknown> {
  const d = await getDigest(token);
  const r = await fetch(`${SP_SITE_URL}/_api/web/lists`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=verbose", "X-RequestDigest": d },
    body: JSON.stringify({ __metadata: { type: "SP.List" }, AllowContentTypes: false, BaseTemplate: baseTemplate, ContentTypesEnabled: false, Title: listTitle, Description: description }),
  });
  if (!r.ok) { const t = await r.text(); throw new Error(`createSpList ${r.status}: ${t}`); }
  await new Promise(r => setTimeout(r, 1500));
  return r.status === 204 ? {} : r.json().catch(() => ({}));
}

export async function listExists(
  token: string,
  listTitle: string
): Promise<boolean> {
  const encodedListTitle = encodeURIComponent(listTitle);
  const url = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodedListTitle}')`;
  try {
    const response = await fetch(url, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Authorization': `Bearer ${token}`,
      },
    });
    return response.ok;
  } catch (e: any) {
    return !e?.response?.ok;
  }
}

// ── Low-level HTTP helpers (from reference) ─────────────────────────────────────
export async function spGet(token: string, url: string): Promise<unknown> {
  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
    },
  });
  if (!response.ok) throw new Error(`GET ${response.status} ${url}`);
  return response.json();
}

export async function spPost(token: string, url: string, body: unknown): Promise<unknown> {
  const digest = await getDigest(token);
  const cleanBody = body ? JSON.parse(JSON.stringify(body)) : {};
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': digest,
    },
    body: JSON.stringify(cleanBody),
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`POST ${response.status}: ${text}`);
  }
  return response.status === 204 ? {} : response.json().catch(() => ({}));
}

export async function spPatch(token: string, url: string, body: unknown): Promise<void> {
  const digest = await getDigest(token);
  const cleanBody = body ? JSON.parse(JSON.stringify(body)) : {};
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'X-RequestDigest': digest,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE',
    },
    body: JSON.stringify(cleanBody),
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`PATCH ${response.status}: ${text}`);
  }
}

export async function spDelete(token: string, url: string): Promise<void> {
  const digest = await getDigest(token);
  await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'X-RequestDigest': digest,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE',
    },
  });
}

// ── Version helpers (from reference) ─────────────────────────────────────────
function parseVersion(v: string): { major: number; minor: number } {
  const [major = 1, minor = 0] = (v || '1.0').split('.').map(Number);
  return { major, minor };
}

function formatVersion({ major, minor }: { major: number; minor: number }): string {
  return `${major}.${minor}`;
}

export function incrementMinor(version: string): string {
  const { major, minor } = parseVersion(version);
  return formatVersion({ major, minor: minor + 1 });
}

export function incrementMajor(version: string): string {
  return formatVersion({ major: parseVersion(version).major + 1, minor: 0 });
}

function compareVersions(a: string, b: string): number {
  const pa = parseVersion(a), pb = parseVersion(b);
  return pa.major !== pb.major ? pa.major - pb.major : pa.minor - pb.minor;
}

export function isVersionGreater(a: string, b: string): boolean {
  return compareVersions(a, b) > 0;
}

// ── Form Config CRUD (from reference) ────────────────────────────────────────
interface FormConfigData {
  Id?: string;
  Title: string;
  FormID?: string;
  NumberOfApprovalLayer?: number;
  Slug?: string;
  CurrentVersion?: string;
  IsPublished?: boolean;
  IsPublic?: boolean;
  ConditionField?: string;
  ApprovalRules?: string;
}

export async function getAllFormConfigs(token: string): Promise<FormConfigData[]> {
  if (!await listExists(token, 'Master Form')) return [];
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$select=Id,Title,FormID,NumberOfApprovalLayer,Slug,CurrentVersion,IsPublished,IsPublic,ConditionField,ApprovalRules&$orderby=Title asc&$top=500`) as { value?: FormConfigData[] };
  return data.value || [];
}

export async function getFormConfigByTitle(token: string, listTitle: string): Promise<FormConfigData | null> {
  if (!await listExists(token, 'Master Form')) return null;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Title eq '${encodeURIComponent(listTitle)}'&$select=Id,Title,FormID,NumberOfApprovalLayer,Slug,CurrentVersion,IsPublished,IsPublic,ConditionField,ApprovalRules&$top=1`) as { value?: FormConfigData[] };
  return data.value?.[0] || null;
}

interface UpsertFormConfigParams {
  formId?: string;
  numLayers?: number;
  slug?: string;
  version?: string;
  isPublished?: boolean;
  isPublic?: boolean;
  conditionField?: string;
  approvalRules?: unknown;
}

export async function upsertFormConfig(
  token: string,
  listTitle: string,
  config: UpsertFormConfigParams
): Promise<string> {
  await ensureListExists(token, 'Master Form');
  const existing = await getFormConfigByTitle(token, listTitle);
  const body: Record<string, unknown> = {
    Title: listTitle,
    FormID: config.formId || '',
    NumberOfApprovalLayer: parseInt(String(config.numLayers), 10) || 0,
    Slug: config.slug || '',
    CurrentVersion: config.version || '1.0',
    IsPublished: config.isPublished ?? true,
    IsPublic: config.isPublic ?? true,
    ConditionField: config.conditionField || '',
    ApprovalRules: config.approvalRules ? JSON.stringify(config.approvalRules) : '',
  };

  if (existing?.Id) {
    await spPatch(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items(${existing.Id})`, body);
    return existing.Id;
  }
  const result = await spPost(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items`, body) as { Id?: string };
  if (!result.Id) throw new Error('upsertFormConfig: POST returned no Id');
  return result.Id;
}

// ── Approvers (from reference) ─────────────────────────────────────────────
interface ApproverLayer {
  email: string;
  name?: string;
}

export async function upsertApprovers(token: string, listTitle: string, layers: ApproverLayer[]): Promise<void> {
  await ensureListExists(token, 'Approvers');
  const existing = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items?$filter=FormTitle eq '${encodeURIComponent(listTitle)}'&$select=Id&$top=500`) as { value?: { Id: string }[] };
  for (const item of existing.value || []) {
    await spDelete(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items(${item.Id})`);
  }
  for (let i = 0; i < layers.length; i++) {
    if (!layers[i]?.email) continue;
    await spPost(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items`, {
      Title: `${listTitle} - Layer ${i + 1}`,
      FormTitle: listTitle,
      LayerNumber: i + 1,
      ApproverEmail: layers[i].email,
      ApproverName: layers[i].name || '',
    });
  }
}

// ── Form Deletion ─────────────────────────────────────────────────────────

/**
 * Deletes all version records for a form from the Web Form Versions list.
 */
export async function deleteFormVersions(token: string, formTitle: string): Promise<number> {
  if (!await listExists(token, 'Web Form Versions')) return 0;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(formTitle)}'&$select=Id&$top=500`) as { value?: { Id: number }[] };
  let count = 0;
  for (const item of data.value || []) {
    await spDelete(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items(${item.Id})`);
    count++;
  }
  return count;
}

/**
 * Deletes all audit log entries for a form from the Form Builder Log list.
 */
export async function deleteFormLogEntries(token: string, formTitle: string): Promise<number> {
  if (!await listExists(token, 'Form Builder Log')) return 0;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Form%20Builder%20Log')/items?$filter=FormTitle eq '${encodeURIComponent(formTitle)}'&$select=Id&$top=500`) as { value?: { Id: number }[] };
  let count = 0;
  for (const item of data.value || []) {
    await spDelete(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Form%20Builder%20Log')/items(${item.Id})`);
    count++;
  }
  return count;
}

/**
 * Deletes all approver records for a form from the Approvers list.
 */
export async function deleteFormApprovers(token: string, formTitle: string): Promise<number> {
  if (!await listExists(token, 'Approvers')) return 0;
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items?$filter=FormTitle eq '${encodeURIComponent(formTitle)}'&$select=Id&$top=500`) as { value?: { Id: number }[] };
  let count = 0;
  for (const item of data.value || []) {
    await spDelete(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items(${item.Id})`);
    count++;
  }
  return count;
}

/**
 * Deletes the form config entry from the Master Form list.
 */
export async function deleteFormConfig(token: string, formId: string): Promise<void> {
  await spDelete(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items(${formId})`);
}

export interface DeleteFormResult {
  configDeleted: boolean;
  versionsDeleted: number;
  logEntriesDeleted: number;
  approversDeleted: number;
}

/**
 * Master delete function — deletes a form and all related rows.
 * Cascades: Web Form Versions → Form Builder Log → Approvers → Master Form.
 * Does NOT delete the form's submission list or response list.
 */
export async function deleteForm(token: string, formTitle: string, formId: string): Promise<DeleteFormResult> {
  const [versionsDeleted, logEntriesDeleted, approversDeleted] = await Promise.all([
    deleteFormVersions(token, formTitle),
    deleteFormLogEntries(token, formTitle),
    deleteFormApprovers(token, formTitle),
  ]);
  await deleteFormConfig(token, formId);
  return { configDeleted: true, versionsDeleted, logEntriesDeleted, approversDeleted };
}

// ── Form Versions (from reference) ────────────────────────────────────────
interface FormVersionRecord {
  Title: string;
  FormTitle: string;
  FormSlug: string;
  FormVersion: string;
  SurveyJSON: string;
  PublishedBy: string;
  PublishedAt: string;
}

export async function getFormVersionHistory(token: string, listTitle: string): Promise<FormVersionRecord[]> {
  if (!await listExists(token, 'Web Form Versions')) return [];
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(listTitle)}'&$select=FormVersion,PublishedAt,PublishedBy,Title&$orderby=PublishedAt desc&$top=100`) as { value?: FormVersionRecord[] };
  return data.value || [];
}

export async function logEvent(
  token: string,
  params: {
    formTitle: string;
    eventType: string;
    changedBy: string;
    before?: unknown;
    after?: unknown;
    summary?: string;
  }
): Promise<void> {
  try {
    await ensureListExists(token, 'Form Builder Log');
    await spPost(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Form%20Builder%20Log')/items`, {
      Title: `${params.formTitle} — ${params.eventType}`,
      FormTitle: params.formTitle,
      EventType: params.eventType,
      ChangedBy: params.changedBy,
      EventSummary: params.summary || '',
      BeforeJSON: params.before ? JSON.stringify(params.before) : '',
      AfterJSON: params.after ? JSON.stringify(params.after) : '',
      EventAt: new Date().toISOString(),
    });
  } catch (e) {
    console.warn('[SP] logEvent failed:', (e as Error).message);
  }
}

// ── Diff helpers (from reference) ─────────────────────────────────────────
export function diffSurveyJson(before: unknown, after: unknown): unknown[] {
  if (!before) return [{ type: 'FORM_CREATED', summary: 'Form created' }];
  const events: unknown[] = [];
  const bF = (before as { pages?: { elements?: unknown[] }[] })?.pages?.[0]?.elements || [];
  const aF = (after as { pages?: { elements?: unknown[] }[] })?.pages?.[0]?.elements || [];
  const bM = Object.fromEntries(bF.map((f: unknown) => [(f as { name?: string }).name, f]));
  const aM = Object.fromEntries(aF.map((f: unknown) => [(f as { name?: string }).name, f]));
  for (const f of aF) {
    const fname = (f as { name?: string }).name;
    if (fname && !bM[fname]) events.push({ type: 'FIELD_ADDED', summary: `Field added: "${fname}"`, before: null, after: f });
  }
  for (const f of bF) {
    const fname = (f as { name?: string }).name;
    if (fname && !aM[fname]) events.push({ type: 'FIELD_REMOVED', summary: `Field removed: "${fname}"`, before: f, after: null });
  }
  for (const f of aF) {
    const fname = (f as { name?: string }).name;
    if (!fname) continue;
    const p = bM[fname];
    if (p && JSON.stringify(p) !== JSON.stringify(f)) events.push({ type: 'FIELD_CHANGED', summary: `Field modified: "${fname}"`, before: p, after: f });
  }
  return events;
}

// ── Bootstrap (from reference) ──────────────────────────────────────────
const LIST_SCHEMAS: Record<string, { t: number; desc: string; cols: { n: string; k: number; ml?: boolean; rt?: boolean }[] }> = {
  'Master Form': { t: 100, desc: 'Form builder configuration', cols: [
    { n: 'FormID', k: 2 }, { n: 'NumberOfApprovalLayer', k: 9 },
    { n: 'Slug', k: 2 }, { n: 'CurrentVersion', k: 2 },
    { n: 'IsPublished', k: 8 }, { n: 'IsPublic', k: 8 },
    { n: 'ConditionField', k: 2 }, { n: 'ApprovalRules', k: 3, ml: true },
  ]},
  'Approvers': { t: 100, desc: 'Approver layers per form', cols: [
    { n: 'FormTitle', k: 2 }, { n: 'LayerNumber', k: 9 },
    { n: 'ApproverEmail', k: 2 }, { n: 'ApproverName', k: 2 },
  ]},
  'Web Form Versions': { t: 100, desc: 'Published form version metadata', cols: [
    { n: 'FormTitle', k: 2 }, { n: 'FormSlug', k: 2 },
    { n: 'FormVersion', k: 2 }, { n: 'SurveyJSON', k: 3, ml: true },
    { n: 'PublishedBy', k: 2 }, { n: 'PublishedAt', k: 4 },
  ]},
  'Form Builder Log': { t: 100, desc: 'Audit log', cols: [
    { n: 'FormTitle', k: 2 }, { n: 'EventType', k: 2 },
    { n: 'ChangedBy', k: 2 }, { n: 'EventSummary', k: 3, ml: true },
    { n: 'BeforeJSON', k: 3, ml: true }, { n: 'AfterJSON', k: 3, ml: true },
    { n: 'EventAt', k: 4 },
  ]},
};

async function ensureListExists(token: string, listTitle: string): Promise<void> {
  const schema = LIST_SCHEMAS[listTitle];
  const exists = await listExists(token, listTitle);
  if (!exists) {
    await createSpList(token, listTitle, schema?.t ?? 100, schema?.desc ?? '');
    await new Promise(r => setTimeout(r, 1000));
  }
  if (schema?.cols) {
    for (const col of schema.cols) {
      await addColumn(token, listTitle, col.n, col.k, !!col.ml, !!col.rt);
    }
  }
}

export async function bootstrapSystemLists(token: string, onLog?: (msg: string, type: string) => void): Promise<void> {
  for (const [title, schema] of Object.entries(LIST_SCHEMAS)) {
    onLog?.(`Checking "${title}"…`, 'info');
    if (!await listExists(token, title)) {
      await createSpList(token, title, schema.t, schema.desc);
      onLog?.('✓ Created', 'ok');
    } else {
      onLog?.('✓ Exists', 'ok');
    }
    for (const col of schema.cols) {
      await addColumn(token, title, col.n, col.k, !!col.ml, !!col.rt);
      onLog?.(`  ✓ ${col.n}`, 'ok');
    }
  }
  onLog?.('Bootstrap complete ✓', 'ok');
}

// ── Get latest form by slug (from reference) ────────────────────────────────
export async function getLatestFormBySlug(token: string, slug: string): Promise<{
  formConfig: FormConfigData;
  surveyJson: unknown;
  meta: unknown;
} | null> {
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Slug eq '${encodeURIComponent(slug)}'&$select=Title,CurrentVersion,FormID,NumberOfApprovalLayer,Slug,IsPublished,IsPublic,ConditionField,ApprovalRules&$top=1`) as { value?: FormConfigData[] };
  const form = data.value?.[0];
  if (!form) return null;
  if (!form.IsPublished) return null;

  const versionData = await getFormVersionByTitle(token, form.Title, form.CurrentVersion || '1.0');
  return {
    formConfig: form,
    surveyJson: versionData?.surveyJson || null,
    meta: versionData?.meta || {},
  };
}

async function getFormVersionByTitle(token: string, listTitle: string, version: string): Promise<{ surveyJson: unknown; meta: unknown } | null> {
  const data = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(listTitle)}' and FormVersion eq '${encodeURIComponent(version)}'&$select=SurveyJSON,FormVersion,PublishedAt,PublishedBy&$top=1`) as { value?: { SurveyJSON?: string }[] };
  const row = data.value?.[0];
  if (!row?.SurveyJSON) return null;
  try {
    return JSON.parse(row.SurveyJSON);
  } catch {
    return null;
  }
}

// ── Response List Provisioning ────────────────────────────────────────────

/**
 * Provisions a dedicated SP list for form responses.
 * Creates [FormTitle] Responses list with system columns + per-field columns.
 * Idempotent — safe to call multiple times.
 */
export async function provisionResponseList(
  token: string,
  formTitle: string,
  surveyJson: unknown,
  onLog: (msg: string, type: string) => void = () => {}
): Promise<void> {
  const listName = `${formTitle} Responses`;
  onLog(`Checking response list "${listName}"…`, 'info');

  // Create list if missing
  if (!(await listExists(token, listName))) {
    await createSpList(token, listName, 100, `Form responses for ${formTitle}`);
    onLog(`Created list "${listName}"`, 'ok');
  } else {
    onLog(`List exists`, 'ok');
  }

  // Always-present system columns
  await addColumn(token, listName, 'SubmittedBy', 2); // Text — email or "anonymous"
  await addColumn(token, listName, 'SubmittedAt', 4); // DateTime
  await addColumn(token, listName, 'Status', 2); // Text — Submitted/Pending/Approved/Rejected
  await addColumn(token, listName, 'CurrentApprovalLayer', 9); // Number
  await addColumn(token, listName, 'FormVersion', 2); // Text
  await addColumn(token, listName, 'RawJSON', 3, true); // Note — full survey.data JSON backup
  onLog('  ✓ System columns', 'ok');

  // Per-field columns from survey JSON
  if (!surveyJson || typeof surveyJson !== 'object') {
    onLog('  ⚠ No survey JSON, skipping field columns', 'warn');
    return;
  }

  const questions = flattenQuestions(surveyJson as SurveyJson);
  for (const q of questions) {
    if (!q.type) continue;

    if (q.type === 'matrixdynamic' || q.type === 'tableinput') {
      // Matrix/table types create two columns: _Html and _Json
      const fieldName = q.name;
      if (fieldName) {
        await addColumn(token, listName, `${fieldName}_Html`, 3, true, true); // Enhanced Rich Text
        await addColumn(token, listName, `${fieldName}_Json`, 3, true, false);
        onLog(`  ✓ ${fieldName}_Html + ${fieldName}_Json`, 'ok');
      }
    } else {
      const kind = getSpColumnKind(q);
      if (!kind) continue; // html, panel — skip
      const fieldName = q.name;
      if (!fieldName) continue;

      // Extract choices for Choice (6) and MultiChoice (15) columns
      let choiceValues: string[] | undefined;
      if (kind.FieldTypeKind === 6 || kind.FieldTypeKind === 15) {
        const src = (q as { spChoicesSource?: { list?: string; column?: string } }).spChoicesSource;
        if (src?.list && src?.column) {
          try {
            choiceValues = await getSharePointChoices(src.list, src.column, token);
            onLog(`  ↳ Fetched ${choiceValues.length} choices from "${src.list}.${src.column}"`, 'info');
          } catch {
            choiceValues = [];
          }
        }
        if (!choiceValues || choiceValues.length === 0) {
          const rawChoices = (q as { choices?: (string | { value: string; text: string })[] }).choices;
          if (Array.isArray(rawChoices) && rawChoices.length > 0) {
            choiceValues = rawChoices.map((c) => (typeof c === 'string' ? c : c.value || c.text || '')).filter(Boolean);
          }
        }
      }

      await addColumn(token, listName, fieldName, kind.FieldTypeKind, kind.FieldTypeKind === 3, false, choiceValues);
      onLog(`  ✓ ${fieldName} (${kind.label})`, 'ok');
    }
  }
  onLog('Provisioning complete ✓', 'ok');
}

// ── Dynamic Matrix → HTML Serialization ────────────────────────────────────

/**
 * Converts a dynamicmatrix response to HTML table for SP storage.
 */
export function dynamicMatrixToHtml(
  rows: unknown,
  questionDef: unknown
): string {
  if (!Array.isArray(rows) || rows.length === 0) {
    return '<em>No rows</em>';
  }

  const qDef = questionDef as { columns?: { title?: string; name?: string }[] };
  const columns = qDef.columns || [];

  // Header
  const headers = ['#', ...columns.map((c) => c.title || c.name)];
  const headerHtml = headers
    .map(
      (h) =>
        `<th style="border:1px solid #ccc;padding:8px;background:#f0f0f0;text-align:left">${h}</th>`
    )
    .join('');

  // Rows
  const bodyHtml = rows
    .map((row: unknown, i: number) => {
      const r = row as Record<string, unknown>;
      const cells = [
        i + 1,
        ...columns.map((c) => {
          const v = r[c.name ?? ''];
          if (Array.isArray(v)) return v.join(', ');
          return v ?? '';
        }),
      ];
      return `<tr>${cells
        .map(
          (c) =>
            `<td style="border:1px solid #ccc;padding:8px;vertical-align:top">${c}</td>`
        )
        .join('')}</tr>`;
    })
    .join('');

  return `<table style="border-collapse:collapse;width:100%;font-family:sans-serif;font-size:13px">
    <thead><tr>${headerHtml}</tr></thead>
    <tbody>${bodyHtml}</tbody>
  </table>`;
}

// ── Email Notifications via SharePoint ─────────────────────────────────────

interface EmailParams {
  to: string | string[];
  subject: string;
  body: string;
}

/**
 * Sends email via SharePoint REST API (_api/SP.Utilities.Utility.SendEmail)
 */
export async function sendSpEmail(token: string, { to, subject, body }: EmailParams): Promise<void> {
  const digest = await getDigest(token);
  const url = `${SP_SITE_URL}/_api/SP.Utilities.Utility.SendEmail`;

  const recipients = Array.isArray(to) ? to : [to];

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'X-RequestDigest': digest,
    },
    body: JSON.stringify({
      properties: {
        __metadata: { type: 'SP.Utilities.EmailProperties' },
        To: { results: recipients },
        Subject: subject,
        Body: body,
      },
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`sendSpEmail failed ${response.status}: ${text}`);
  }
}

// ── Approval Notification Triggers ─────────────────────────────────────────

interface ApprovalNotificationParams {
  formTitle: string;
  submittedBy: string;
  responseItemId: number;
  layer: number;
  totalLayers: number;
  action?: 'submit' | 'approve' | 'reject';
  nextApproverEmail?: string;
}

/**
 * Triggers email notifications for approval workflow.
 * Handles: new submission, layer approved, final approval, rejection.
 */
export async function triggerApprovalNotification(
  token: string,
  params: ApprovalNotificationParams
): Promise<void> {
  const { formTitle, submittedBy, responseItemId, layer, totalLayers, action = 'submit', nextApproverEmail } = params;

  try {
    if (action === 'submit') {
      // New submission - notify layer 1 approver
      const approvers = await spGet(
        token,
        `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items?$filter=FormTitle eq '${encodeURIComponent(formTitle)}' and LayerNumber eq ${layer}&$select=ApproverEmail,ApproverName&$top=1`
      ) as { value?: { ApproverEmail?: string; ApproverName?: string }[] };

      const approver = approvers.value?.[0];
      if (approver?.ApproverEmail) {
        await sendSpEmail(token, {
          to: approver.ApproverEmail,
          subject: `[Action Required] New ${formTitle} submission by ${submittedBy}`,
          body: `<p>A new submission for <strong>${formTitle}</strong> requires your approval.</p>
<p><strong>Submitted by:</strong> ${submittedBy}</p>
<p><strong>Submission ID:</strong> ${responseItemId}</p>
<p><strong>Approval Layer:</strong> ${layer} of ${totalLayers}</p>
<p><a href="${SP_SITE_URL}/admin/approvals?form=${encodeURIComponent(formTitle)}&item=${responseItemId}">Review and Approve</a></p>`,
        });
      }
    } else if (action === 'approve') {
      if (layer < totalLayers && nextApproverEmail) {
        // Notify next layer approver
        await sendSpEmail(token, {
          to: nextApproverEmail,
          subject: `[Action Required] ${formTitle} — pending your approval (Layer ${layer + 1})`,
          body: `<p>A submission for <strong>${formTitle}</strong> has been approved at Layer ${layer}.</p>
<p>It now requires your approval at Layer ${layer + 1}.</p>
<p><strong>Submitted by:</strong> ${submittedBy}</p>
<p><strong>Submission ID:</strong> ${responseItemId}</p>
<p><a href="${SP_SITE_URL}/admin/approvals?form=${encodeURIComponent(formTitle)}&item=${responseItemId}">Review and Approve</a></p>`,
        });
      } else if (layer === totalLayers) {
        // Final approval - notify submitter
        await sendSpEmail(token, {
          to: submittedBy,
          subject: `[Approved] Your ${formTitle} submission has been approved`,
          body: `<p>Your submission for <strong>${formTitle}</strong> has been fully approved.</p>
<p><strong>Submission ID:</strong> ${responseItemId}</p>
<p>All approval layers have completed. Thank you!</p>`,
        });
      }
    } else if (action === 'reject') {
      // Notify submitter of rejection
      await sendSpEmail(token, {
        to: submittedBy,
        subject: `[Rejected] Your ${formTitle} submission was not approved`,
        body: `<p>Your submission for <strong>${formTitle}</strong> was not approved.</p>
<p><strong>Submission ID:</strong> ${responseItemId}</p>
<p>Please contact your administrator for more details.</p>`,
      });
    }
  } catch (e) {
    console.warn('[triggerApprovalNotification] failed:', (e as Error).message);
    // Don't throw - email failures shouldn't block the workflow
  }
}

// ── Signature Image Upload ─────────────────────────────────────────────
// Signatures are uploaded as PNG files to a "Signature Images" document
// library and linked via a URL/Hyperlink column in the response list.
//
// File naming: {action}-{formId}-{yymmdd}{xxx}.png
//   action  = "submission" | "approval" | "reject"
//   formId  = form identifier
//   yymmdd  = local date (2-digit year, 2-digit month, 2-digit day)
//   xxx     = daily counter starting at 001

const SIGNATURE_LIBRARY = "Signature Images";

/** Get the next daily counter by checking existing files for today */
async function getNextSignatureCounter(token: string, formId: string, action: string): Promise<string> {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const prefix = `${action}-${formId}-${yy}${mm}${dd}`;

  try {
    // List files matching today's prefix
    const query = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(SIGNATURE_LIBRARY)}')/rootfolder/files?$select=Name&$filter=startswith(Name,'${encodeURIComponent(prefix)}')&$orderby=Name desc&$top=1`;
    const data = await spGet(token, query) as { value?: { Name?: string }[] };

    const lastName = data.value?.[0]?.Name;
    if (lastName) {
      const match = lastName.match(/^.+(\d{3})\.png$/);
      if (match) {
        return String(parseInt(match[1], 10) + 1).padStart(3, '0');
      }
    }
  } catch {
    // Library might not exist yet — start at 001
  }

  return '001';
}

/**
 * Upload a base64 signature image to the Signature Images document library.
 * Returns the server-relative URL to the uploaded file.
 */
export async function uploadSignatureImage(
  token: string,
  formId: string,
  action: "submission",
  base64DataUrl: string,
): Promise<string> {
  // Strip the data URI prefix and convert to binary
  const base64 = base64DataUrl.replace(/^data:image\/\w+;base64,/, '');
  const binaryString = atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }

  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const counter = await getNextSignatureCounter(token, formId, action);
  const fileName = `${action}-${formId}-${yy}${mm}${dd}${counter}.png`;

  const sitePath = new URL(SP_SITE_URL).pathname;
  const result = await spUploadFile(token, SIGNATURE_LIBRARY, fileName, bytes) as { ServerRelativeUrl?: string };
  return result.ServerRelativeUrl ?? `${sitePath}/${SIGNATURE_LIBRARY}/${fileName}`;
}
