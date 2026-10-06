import * as React from 'react';
import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './FileUploaed.module.scss';
import type { IFileUploaedProps } from './IFileUploaedProps';

const FILE_PREFIX   = 'BulkUploadTracker';
const SAVED_LIBRARY = 'ComplianceBulkUploads';

// ─── Data model ────────────────────────────────────────────────────────────────

type RowKind = 'section' | 'subitem' | 'continuation' | 'grouplabel';

interface IChecklistRow {
  kind: RowKind;
  sno: string;
  particulars: string;
  department: string;
  documentLocation: string;
  member: string;
  status: string;
  remarks: string;
  compliedAsOn: string;
  extra?: Record<string, any>
}

interface ISheetData {
  name: string;
  spvName: string;
  date: string;
  rows: IChecklistRow[];
}

interface IUploadGroup {
  fileName: string;
  label: string;
  uploadedAt: Date;
  sheets: ISheetData[];
  collapsed: boolean;
}

interface ISharePointFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified: string;
  Length: string;
}

interface IFileUploaedState {
  uploads: IUploadGroup[];
  activeSheetIdx: Record<number, number>;
  configMap: Record<string, string>;
  configLoading: boolean;
  configError: string;
  error: string;
  savingExcel: boolean;
  saveError: string;
  modalOpen: boolean;
  modalTitle: string;
  modalFolderUrl: string;
  modalFiles: ISharePointFile[];
  modalLoading: boolean;
  modalError: string;
  uploadingFile: boolean;
  uploadError: string;
  uploadSuccess: string;
}

// ─── Status colours ─────────────────────────────────────────────────────────────

const STATUS_META: Record<string, { bg: string; color: string }> = {
  'complied':        { bg: '#e6f4ec', color: '#1a7a3c' },
  'completed':       { bg: '#e6f4ec', color: '#1a7a3c' },
  'pending':         { bg: '#fff4e5', color: '#b45309' },
  'to be complied':  { bg: '#fff4e5', color: '#b45309' },
  'not applicable':  { bg: '#f3f4f6', color: '#6b7280' },
  'in review':       { bg: '#eff6ff', color: '#1d4ed8' },
  'deferred':        { bg: '#fdf4ff', color: '#7c3aed' },
  'rejected':        { bg: '#fee2e2', color: '#b91c1c' },
};

function statusStyle(status: string): React.CSSProperties {
  const key = status.trim().toLowerCase();
  const meta = STATUS_META[key] || { bg: '#f3f4f6', color: '#374151' };
  return {
    display: 'inline-block',
    padding: '3px 10px',
    borderRadius: 20,
    background: meta.bg,
    color: meta.color,
    fontWeight: 600,
    fontSize: 11,
    whiteSpace: 'nowrap' as const,
    letterSpacing: '0.2px',
  };
}

// ─── Utility helpers ─────────────────────────────────────────────────────────────

function formatBytes(bytes: string): string {
  const n = parseInt(bytes, 10);
  if (isNaN(n)) return '—';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

function normaliseDate(raw: string): string {
  if (!raw || raw === '-') return '';
  const num = Number(raw);
  if (!isNaN(num) && num > 1000) {
    return new Date(Math.round((num - 25569) * 86400 * 1000))
      .toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  return raw;
}

function serialToDate(serial: number): string {
  return new Date(Math.round((serial - 25569) * 86400 * 1000))
    .toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
}

function toServerRelativePath(url: string, siteUrl: string): string {
  if (!url) return '';
  if (url.indexOf('/') === 0 && url.indexOf('//') !== 0) return url;
  try { return new URL(url).pathname; } catch { /**/ }
  try {
    const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
    return `${sitePath}/${url.replace(/^\//, '')}`;
  } catch { return url; }
}

function buildTimestamp(d: Date): string {
  const pad = (n: number): string => n < 10 ? '0' + n : String(n);
  return `_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function buildLabel(d: Date): string {
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) +
    ', ' + d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
}

function parseDateFromFileName(name: string): Date | undefined {
  const m = name.match(/(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})/);
  if (!m) return undefined;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6]));
}

// ─── Excel parser ──────────────────────────────────────────────────────────────

function classifyRow(raw: unknown[]): RowKind {
  const sno    = String(raw[0] || '').trim();
  const part   = String(raw[1] || '').trim();
  const dept   = String(raw[2] || '').trim();
  const member = String(raw[4] || '').trim();
  const status = String(raw[5] || '').trim();

  if (!sno && !dept && !member && !status) {
    return part ? 'grouplabel' : 'continuation';
  }
  if (!sno && (dept || member || status)) return 'continuation';
  const num = Number(sno);
  if (!isNaN(num) && num > 0 && !dept && !member && !status) return 'section';
  return 'subitem';
}

function parseSheetRows(ws: XLSX.WorkSheet): { rows: IChecklistRow[]; spvName: string; date: string } {
  const raw: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][];

  let spvName = '';
  let date    = '';
  let headerIdx = -1;

  for (let i = 0; i < Math.min(raw.length, 10); i++) {
    const r = raw[i];
    const c0 = String(r[0] || '').trim().toLowerCase();
    const c1 = String(r[1] || '').trim();

    if (c0 === 'spv')  { spvName = c1; continue; }
    if (c0 === 'date') { date = c1 && Number(c1) > 1000 ? serialToDate(Number(c1)) : c1; continue; }
    if (c0 === 'sr. no.' || c0 === 'sr.no.' || c0 === 'sr no') {
      headerIdx = i;
      break;
    }
  }

  const headers = (headerIdx >= 0 && raw[headerIdx]) ? raw[headerIdx] : [];
  const dataRows = headerIdx >= 0 ? raw.slice(headerIdx + 1) : raw;

  const rows: IChecklistRow[] = [];

  for (const r of dataRows) {

    const sno     = String(r[0] || '').trim();
    const part    = String(r[1] || '').trim();
    const dept    = String(r[2] || '').trim();
    const docLoc  = String(r[3] || '').trim();
    const member  = String(r[4] || '').trim();
    const status  = String(r[5] || '').trim();
    const remarks = String(r[6] || '').trim();
    const dateRaw = String(r[7] || '').trim();

    if (!sno && !part && !dept && !member && !status && !remarks) continue;

    const kind = classifyRow(r);

    // ✅ Dynamic column capture
    const extra: Record<string, any> = {};

    if (headers && headers.length > 8) {
      headers.forEach((h: any, idx: number) => {
        if (idx > 7 && h) {
          extra[String(h).trim()] = r[idx];
        }
      });
    }

    rows.push({
      kind,
      sno,
      particulars: part,
      department: dept,
      documentLocation: docLoc,
      member,
      status,
      remarks,
      compliedAsOn: normaliseDate(dateRaw),
      extra // ✅ safe add
    });
  }

  return { rows, spvName, date };
}

function parseWorkbook(buffer: ArrayBuffer): ISheetData[] {
  const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
  return wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const { rows, spvName, date } = parseSheetRows(ws);
    return { name, rows, spvName, date };
  }).filter(s => s.rows.length > 0);
}

// ─── Component ────────────────────────────────────────────────────────────────

export default class FileUploaed extends React.Component<IFileUploaedProps, IFileUploaedState> {
  constructor(props: IFileUploaedProps) {
    super(props);
    this.state = {
      uploads: [],
      activeSheetIdx: {},
      configMap: {},
      configLoading: true,
      configError: '',
      error: '',
      savingExcel: false,
      saveError: '',
      modalOpen: false,
      modalTitle: '',
      modalFolderUrl: '',
      modalFiles: [],
      modalLoading: false,
      modalError: '',
      uploadingFile: false,
      uploadError: '',
      uploadSuccess: '',
    };
  }

  public componentDidMount(): void {
    this._loadConfigurationList();
    this._loadAllSavedExcels();
  }

  private _loadConfigurationList(): void {
    const { spHttpClient, siteUrl } = this.props;
    const url = `${siteUrl}/_api/web/lists/getbytitle('ConfigurationList')/items?$select=MappingField,LocationURL&$top=5000`;
    spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: { MappingField: string; LocationURL: string }[] }) => {
        const configMap: Record<string, string> = {};
        (data.value || []).forEach(item => {
          if (item.MappingField) configMap[String(item.MappingField).trim()] = item.LocationURL || '';
        });
        this.setState({ configMap, configLoading: false });
      })
      .catch(() => this.setState({ configError: 'Could not load ConfigurationList from SharePoint.', configLoading: false }));
  }

private _findMatchingUpload(): IUploadGroup | undefined {
  if (this.state.uploads.length > 0) {
    return this.state.uploads[0]; // always update latest uploaded file
  }
  return undefined;
}

  private _loadAllSavedExcels(): void {
    const { spHttpClient, siteUrl } = this.props;
    const sitePath = toServerRelativePath(siteUrl, siteUrl);
    const folderUrl = `${sitePath}/${SAVED_LIBRARY}`;
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files?$select=Name,ServerRelativeUrl&$orderby=Name`;

    spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => { if (!res.ok) return null; return res.json(); })
      .then((data: { value: { Name: string; ServerRelativeUrl: string }[] } | null) => {
        if (!data || !data.value) return;
        const trackerFiles = data.value.filter(f =>
          f.Name.indexOf(FILE_PREFIX) === 0 &&
          (f.Name.indexOf('.xlsx') !== -1 || f.Name.indexOf('.xls') !== -1)
        );
        if (trackerFiles.length === 0) return;

        const fetches = trackerFiles.map(f => {
          const fileApiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(f.ServerRelativeUrl)}')/$value`;
          return spHttpClient.get(fileApiUrl, SPHttpClient.configurations.v1)
            .then((res: SPHttpClientResponse) => { if (!res.ok) return null; return res.arrayBuffer(); })
            .then((buffer: ArrayBuffer | null) => {
              if (!buffer) return null;
              try {
                const sheets = parseWorkbook(buffer);
                if (sheets.length === 0) return null;
                const uploadedAt = parseDateFromFileName(f.Name) || new Date();
                return { fileName: f.Name, label: buildLabel(uploadedAt), uploadedAt, sheets, collapsed: false } as IUploadGroup;
              } catch { return null; }
            });
        });

        Promise.all(fetches).then(results => {
          const uploads = (results.filter(r => r !== null) as IUploadGroup[])
            .sort((a, b) => b.uploadedAt.getTime() - a.uploadedAt.getTime());
          if (uploads.length > 0) this.setState({ uploads });
        }).catch((): void => undefined);
      })
      .catch(() => undefined);
  }

  private _saveExcelToLibrary(arrayBuffer: ArrayBuffer, fileName: string): void {
    const { spHttpClient, siteUrl } = this.props;
    const sitePath  = toServerRelativePath(siteUrl, siteUrl);
    const folderUrl = `${sitePath}/${SAVED_LIBRARY}`;
    const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files/add(overwrite=true,url='${encodeURIComponent(fileName)}')`;

    this.setState({ savingExcel: true, saveError: '' });
    spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
      body: arrayBuffer as unknown as string,
    })
    .then((res: SPHttpClientResponse) => {
      if (res.ok) { this.setState({ savingExcel: false }); }
      else return res.json().then((err: { error?: { message?: { value?: string } } }) => {
        this.setState({ savingExcel: false, saveError: (err.error?.message?.value) || 'Could not save file.' });
      });
    })
    .catch(() => this.setState({ savingExcel: false, saveError: 'Could not save file to SharePoint library.' }));
  }

 private handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = (evt): void => {
    const buffer = (evt.target as FileReader).result as ArrayBuffer;

    try {
      const sheets = parseWorkbook(buffer);
      const now = new Date();

      // ✅ Check if same file already exists
      const existing = this._findMatchingUpload();

      const savedFileName = existing
        ? existing.fileName
        : FILE_PREFIX + buildTimestamp(now) + '.xlsx';

      if (existing) {

        // ✅ UPDATE EXISTING GROUP
        const uploads = this.state.uploads.slice();
        let index = -1;

        for (let i = 0; i < uploads.length; i++) {
          if (uploads[i].fileName === existing.fileName) {
            index = i;
            break;
          }
        }

        if (index !== -1) {
          const target = uploads[index];
          const mergedSheets: ISheetData[] = target.sheets.slice();

          for (let i = 0; i < sheets.length; i++) {
            const newSheet = sheets[i];
            let found: ISheetData | undefined = undefined;

            for (let j = 0; j < mergedSheets.length; j++) {
              if (mergedSheets[j].name === newSheet.name) {
                found = mergedSheets[j];
                break;
              }
            }

            if (found) {
              // ✅ Append rows safely
              found.rows = found.rows.concat(newSheet.rows);
            } else {
              // ✅ Add new sheet
              mergedSheets.push(newSheet);
            }
          }

          uploads[index] = {
            ...target,
            sheets: mergedSheets,
            uploadedAt: now,
            label: buildLabel(now)
          };
        }

        this.setState({
          uploads: uploads,
          error: ''
        });

      } else {

        // ✅ CREATE NEW GROUP
        const newGroup: IUploadGroup = {
          fileName: savedFileName,
          label: buildLabel(now),
          uploadedAt: now,
          sheets: sheets,
          collapsed: false
        };

        this.setState(prev => ({
          uploads: [newGroup].concat(prev.uploads),
          error: ''
        }));
      }

      // ✅ SharePoint save (overwrite works automatically)
      this._saveExcelToLibrary(buffer, savedFileName);

    } catch {
      this.setState({
        error: 'Failed to parse Excel file. Please check the file format.'
      });
    }
  };

  reader.readAsArrayBuffer(file);
  e.target.value = '';
};

  private _toggleGroup = (idx: number): void => {
    const uploads = this.state.uploads.map((u, i) => i === idx ? { ...u, collapsed: !u.collapsed } : u);
    this.setState({ uploads });
  };

  private _setActiveSheet = (groupIdx: number, sheetIdx: number): void => {
    this.setState(prev => ({ activeSheetIdx: { ...prev.activeSheetIdx, [groupIdx]: sheetIdx } }));
  };

  private _openFolderModal = (sno: string, folderUrl: string): void => {
    this.setState(
      { modalOpen: true, modalTitle: `Files for ${sno}`, modalFolderUrl: folderUrl, modalFiles: [], modalLoading: true, modalError: '' },
      () => this._loadFolderFiles(folderUrl)
    );
  };

  private _loadFolderFiles(folderUrl: string): void {
    const { spHttpClient, siteUrl } = this.props;
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(toServerRelativePath(folderUrl, siteUrl))}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=Name`;
    spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: ISharePointFile[] }) => {
        if (data.value) this.setState({ modalFiles: data.value, modalLoading: false });
        else this.setState({ modalError: 'No files found or folder does not exist.', modalLoading: false });
      })
      .catch(() => this.setState({ modalError: 'Failed to load files from SharePoint folder.', modalLoading: false }));
  }

  private _closeModal = (): void => {
    this.setState({ modalOpen: false, modalFiles: [], modalError: '', uploadError: '', uploadSuccess: '' });
  };

  private _handleUpload = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    e.target.value = '';
    const { spHttpClient, siteUrl } = this.props;
    const { modalFolderUrl } = this.state;
    const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(toServerRelativePath(modalFolderUrl, siteUrl))}')/Files/add(overwrite=true,url='${encodeURIComponent(file.name)}')`;
    this.setState({ uploadingFile: true, uploadError: '', uploadSuccess: '' });
    const reader = new FileReader();
    reader.onload = (evt): void => {
      spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
        body: (evt.target as FileReader).result as unknown as string,
      })
      .then((res: SPHttpClientResponse) => {
        if (res.ok) {
          this.setState({ uploadingFile: false, uploadSuccess: `"${file.name}" uploaded successfully.` });
          this._loadFolderFiles(modalFolderUrl);
        } else {
          return res.json().then((err: { error?: { message?: { value?: string } } }) =>
            this.setState({ uploadingFile: false, uploadError: err.error?.message?.value || 'Upload failed.' })
          );
        }
      })
      .catch(() => this.setState({ uploadingFile: false, uploadError: 'Upload failed. Please try again.' }));
    };
    reader.readAsArrayBuffer(file);
  };

  // ─── Render helpers ────────────────────────────────────────────────────────────

  private _renderSheetTable(sheet: ISheetData): React.ReactElement {
    const { configMap } = this.state;
    const rows = sheet.rows;

    return (
      <div className={styles.sheetTableWrap}>
        {(sheet.spvName || sheet.date) && (
          <div className={styles.sheetMeta}>
            {sheet.spvName && <span className={styles.sheetMetaSpv}><strong>SPV:</strong> {sheet.spvName}</span>}
            {sheet.date    && <span className={styles.sheetMetaDate}><strong>Date:</strong> {sheet.date}</span>}
          </div>
        )}
        <table className={styles.checklistTable}>
          <thead>
            <tr>
              <th className={styles.colSno}>Sr. No.</th>
              <th className={styles.colParticulars}>Particulars</th>
              <th className={styles.colDept}>Department</th>
              <th className={styles.colMember}>Member</th>
              <th className={styles.colStatus}>Status</th>
              <th className={styles.colRemarks}>Remarks</th>
              <th className={styles.colDate}>Complied As On</th>
              <th className={styles.colDocs}>Attachment</th>
              {rows.length > 0 && rows[0].extra &&
                 Object.keys(rows[0].extra).map((key, i) => (
                   <th key={i}>{key}</th>
                      ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              if (row.kind === 'section') {
                return (
                  <tr key={i} className={styles.sectionRow}>
                    <td colSpan={8}>
                      <span className={styles.sectionNum}>{row.sno}.</span>
                      <span className={styles.sectionTitle}>{row.particulars}</span>
                    </td>
                  </tr>
                );
              }

              if (row.kind === 'grouplabel') {
                return (
                  <tr key={i} className={styles.groupLabelRow}>
                    <td colSpan={8}>{row.particulars}</td>
                  </tr>
                );
              }

              const isContinuation = row.kind === 'continuation';
              const spUrl = configMap[row.sno.trim()] || '';
              const trClass = isContinuation ? styles.continuationRow : (i % 2 === 0 ? styles.rowEven : styles.rowOdd);

              return (
                <tr key={i} className={trClass}>
                  <td className={styles.snoCell}>{isContinuation ? '' : row.sno}</td>
                  <td className={styles.particularsCell}>
                    {!isContinuation && row.particulars
                      ? <span>{row.particulars}</span>
                      : isContinuation ? null : <span className={styles.noData}>—</span>}
                  </td>
                  <td>
                    {row.department
                      ? <span className={styles.deptChip}>{row.department}</span>
                      : <span className={styles.noData}>—</span>}
                  </td>
                  <td className={styles.memberCell}>
                    {row.member || <span className={styles.noData}>—</span>}
                  </td>
                  <td>
                    <span
                      style={statusStyle(row.status)}
                      className={styles.statusBadgeClickable}
                      onClick={() => {
                        const options = [
                          'Complied',
                          'Pending',
                          'Not Applicable',
                          'Deferred'
                        ];

                        const currentIndex = options.indexOf(row.status);
                        const nextStatus = options[(currentIndex + 1) % options.length];

                        const uploads = [...this.state.uploads];

                        uploads.forEach(group => {
                          group.sheets.forEach(sheet => {
                            sheet.rows.forEach(r => {
                              if (r === row) {
                                r.status = nextStatus;
                              }
                            });
                          });
                        });

                        this.setState({ uploads });
                      }}
                    >
                      {row.status || 'Select'}
                    </span>
                  </td>
                  <td className={styles.remarksCell}>
                    {row.remarks && row.remarks !== '-'
                      ? row.remarks
                      : <span className={styles.noData}>—</span>}
                  </td>
                  <td className={styles.dateCell}>
                    {row.compliedAsOn || <span className={styles.noData}>—</span>}
                  </td>
                  <td>
                    {row.documentLocation && row.documentLocation.toLowerCase() !== 'document link'
                      ? <a href={row.documentLocation} target="_blank" rel="noreferrer" className={styles.docLink}>↗ View</a>
                      : spUrl
                        ? <button className={styles.spLinkBtn} onClick={() => this._openFolderModal(row.sno || row.particulars.slice(0, 30), spUrl)}>📁 Files</button>
                        : <span className={styles.noData}>—</span>}
                  </td>
                  {/* ✅ Dynamic Columns Cells (SAFE ADDITION) */}
                    {row.extra &&
                      Object.keys(row.extra).map((key, k) => (
                        <td key={k}>
                          {row.extra && row.extra[key] || <span className={styles.noData}>—</span>}
                        </td>
                      ))
                    }
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }

  private _renderGroup(group: IUploadGroup, groupIdx: number): React.ReactElement {
    const activeSheet = this.state.activeSheetIdx[groupIdx] || 0;
    const totalRows = group.sheets.reduce((s, sh) => s + sh.rows.filter(r => r.kind === 'subitem' || r.kind === 'continuation').length, 0);

    return (
      <div key={groupIdx} className={styles.uploadGroup}>
        <div className={styles.groupHeader} onClick={() => this._toggleGroup(groupIdx)}>
          <div className={styles.groupHeaderLeft}>
            <span className={styles.groupToggle}>{group.collapsed ? '▶' : '▼'}</span>
            <span className={styles.groupIcon}>📋</span>
            <div>
              <span className={styles.groupLabel}>Uploaded: {group.label}</span>
              <span className={styles.groupFileName}>{group.fileName}</span>
            </div>
          </div>
          <div className={styles.groupHeaderRight}>
            <span className={styles.groupBadge}>{group.sheets.length} sheet{group.sheets.length !== 1 ? 's' : ''}</span>
            <span className={styles.groupBadgeSecondary}>{totalRows} item{totalRows !== 1 ? 's' : ''}</span>
          </div>
        </div>

        {!group.collapsed && (
          <div className={styles.groupBody}>
            <div className={styles.sheetTabs}>
              {group.sheets.map((sheet, si) => (
                <button
                  key={si}
                  className={`${styles.sheetTab} ${si === activeSheet ? styles.sheetTabActive : ''}`}
                  onClick={() => this._setActiveSheet(groupIdx, si)}
                >
                  {sheet.name}
                </button>
              ))}
            </div>
            {group.sheets[activeSheet] && this._renderSheetTable(group.sheets[activeSheet])}
          </div>
        )}
      </div>
    );
  }

  private _renderModal(): React.ReactElement | null {
    const { modalOpen, modalTitle, modalFiles, modalLoading, modalError, uploadingFile, uploadError, uploadSuccess } = this.state;
    if (!modalOpen) return null;
    return (
      <div className={styles.modalOverlay} onClick={this._closeModal}>
        <div className={styles.modalBox} onClick={e => e.stopPropagation()}>
          <div className={styles.modalHeader}>
            <div className={styles.modalHeaderLeft}>
              <div className={styles.modalFolderIcon}>📁</div>
              <div><div className={styles.modalTitle}>{modalTitle}</div></div>
            </div>
            <button className={styles.modalClose} onClick={this._closeModal} title="Close">✕</button>
          </div>
          <div className={styles.modalBody}>
            <div className={styles.uploadSection}>
              <label className={styles.uploadSectionLabel} htmlFor="modalFileUpload">
                {uploadingFile ? <><div className={styles.spinnerSm} /> Uploading...</> : <>⬆ Upload File to this Folder</>}
              </label>
              <input id="modalFileUpload" type="file" className={styles.fileInput} onChange={this._handleUpload} disabled={uploadingFile} />
            </div>
            {uploadError   && <div className={styles.errorMsg}>{uploadError}</div>}
            {uploadSuccess && <div className={styles.successMsg}>{uploadSuccess}</div>}
            <div className={styles.modalDivider} />
            {modalLoading && <div className={styles.modalSpinner}><div className={styles.spinner} /><span>Loading files...</span></div>}
            {!modalLoading && modalError && <div className={styles.errorMsg}>{modalError}</div>}
            {!modalLoading && !modalError && modalFiles.length === 0 && <div className={styles.modalEmpty}>No files found in this folder.</div>}
            {!modalLoading && !modalError && modalFiles.length > 0 && (
              <>
                <div className={styles.modalFileCount}>{modalFiles.length} file{modalFiles.length !== 1 ? 's' : ''} in folder</div>
                <table className={styles.modalTable}>
                  <thead><tr><th>#</th><th>File Name</th><th>Size</th><th>Last Modified</th><th>Actions</th></tr></thead>
                  <tbody>
                    {modalFiles.map((f, idx) => {
                      const absUrl = this._toAbsoluteUrl(f.ServerRelativeUrl);
                      return (
                        <tr key={idx} className={idx % 2 === 0 ? styles.rowEven : styles.rowOdd}>
                          <td className={styles.indexCell}>{idx + 1}</td>
                          <td className={styles.fileNameCell}><span className={styles.fileIcon}>{this._fileIcon(f.Name)}</span>{f.Name}</td>
                          <td>{formatBytes(f.Length)}</td>
                          <td>{formatDate(f.TimeLastModified)}</td>
                          <td>
                            <div className={styles.actionBtns}>
                              <a href={this._viewerUrl(f.Name, absUrl)} target="_blank" rel="noreferrer" className={styles.btnView}>👁 View</a>
                              <a href={absUrl} download={f.Name} className={styles.btnDownload}>⬇ Download</a>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </>
            )}
          </div>
        </div>
      </div>
    );
  }

  private _viewerUrl(name: string, absUrl: string): string {
    const ext = name.split('.').pop()?.toLowerCase() || '';
    return ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].indexOf(ext) !== -1 ? absUrl + '?web=1' : absUrl;
  }

  private _toAbsoluteUrl(serverRelativeUrl: string): string {
    try { return new URL(this.props.siteUrl).origin + serverRelativeUrl; } catch { return serverRelativeUrl; }
  }

  private _fileIcon(name: string): string {
    const ext = name.split('.').pop()?.toLowerCase() || '';
    return ({ pdf: '📕', doc: '📘', docx: '📘', xls: '📗', xlsx: '📗', ppt: '📙', pptx: '📙', png: '🖼', jpg: '🖼', jpeg: '🖼', gif: '🖼', zip: '🗜', txt: '📄', csv: '📊' } as Record<string, string>)[ext] || '📄';
  }

  public render(): React.ReactElement<IFileUploaedProps> {
    const { uploads, configLoading, configError, error, savingExcel, saveError } = this.state;
    const totalItems = uploads.reduce((s, u) =>
      s + u.sheets.reduce((ss, sh) => ss + sh.rows.filter(r => r.kind === 'subitem' || r.kind === 'continuation').length, 0), 0);

    return (
      <section className={styles.fileUploaed}>

        <div className={styles.header}>
          <div className={styles.headerInner}>
            <div className={styles.headerLeft}>
              <img src={require('../assets/dmi-logo.svg')} alt="DMI Finance" className={styles.brandLogo} />
              <div className={styles.headerDivider} />
              <div>
                <h2 className={styles.title}>Bulk Upload Tracker</h2>
                <p className={styles.subtitle}>Enterprise Document Management System</p>
              </div>
            </div>
            <span className={styles.headerBadge}>Internal Tool</span>
          </div>
        </div>

        {configLoading && <div className={styles.infoMsg}>⏳ Loading configuration from SharePoint...</div>}
        {configError   && <div className={styles.errorMsg}>{configError}</div>}
        {saveError     && <div className={styles.errorMsg}>⚠ Could not save to library: {saveError}</div>}

        <div className={styles.toolbar}>
          <div className={styles.toolbarLeft}>
            <label className={styles.uploadLabel} htmlFor="excelUpload">
              <span className={styles.uploadIcon}>📊</span>
              <span>{savingExcel ? '⏳ Saving...' : 'Import Excel File'}</span>
            </label>
            <input id="excelUpload" type="file" accept=".xlsx,.xls" onChange={this.handleFileChange} className={styles.fileInput} disabled={savingExcel} />
            {uploads.length > 0 && (
              <span className={styles.sourceChip}>☁ {uploads.length} file{uploads.length !== 1 ? 's' : ''} loaded</span>
            )}
          </div>
          {totalItems > 0 && (
            <span className={styles.recordBadge}>
              <span className={styles.recordDot} />
              {totalItems} Total Item{totalItems !== 1 ? 's' : ''}
            </span>
          )}
        </div>

        {error && <div className={styles.errorMsg}>{error}</div>}

        {uploads.length > 0 && (
          <div className={styles.groupsWrapper}>
            {uploads.map((group, idx) => this._renderGroup(group, idx))}
          </div>
        )}

        {uploads.length === 0 && !error && (
          <div className={styles.emptyState}>
            <div className={styles.emptyIconWrap}>📋</div>
            <p className={styles.emptyTitle}>No Checklist Loaded</p>
            <p className={styles.emptyDesc}>
              Import an Excel checklist file using the button above.<br />
              Supports multi-sheet checklists with sections, sub-items, and multi-member rows.
            </p>
          </div>
        )}

        {this._renderModal()}
      </section>
    );
  }
}
