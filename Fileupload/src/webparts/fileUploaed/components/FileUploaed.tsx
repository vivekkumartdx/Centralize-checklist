import * as React from 'react';
import * as XLSX from 'xlsx';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './FileUploaed.module.scss';
import type { IFileUploaedProps } from './IFileUploaedProps';

// Fixed name used for the persisted file — always overwritten, never duplicated
const SAVED_FILE_NAME = 'BulkUploadTracker.xlsx';
const SAVED_LIBRARY   = 'ComplianceBulkUploads';

interface IExcelRow {
  SNo: string;
  Particulars: string;
  Department: string;
  DocumentLink: string;
  Member: string;
  Status: string;
  Remarks: string;
  CompliedAsOn: string;
}

interface ISharePointFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified: string;
  Length: string;
}

interface IFileUploaedState {
  rows: IExcelRow[];
  configMap: Record<string, string>;
  configLoading: boolean;
  configError: string;
  error: string;
  fileName: string;
  // 'sharepoint' = loaded from saved file | 'local' = just imported | '' = nothing yet
  dataSource: 'sharepoint' | 'local' | '';
  savingExcel: boolean;
  saveError: string;
  // modal
  modalOpen: boolean;
  modalTitle: string;
  modalFolderUrl: string;
  modalFiles: ISharePointFile[];
  modalLoading: boolean;
  modalError: string;
  // upload within modal
  uploadingFile: boolean;
  uploadError: string;
  uploadSuccess: string;
}

const STATUS_COLORS: Record<string, string> = {
  approved:    '#107c10',
  pending:     '#d83b01',
  'in review': '#0078d4',
  completed:   '#107c10',
  rejected:    '#a4262c',
};

function statusBadgeStyle(status: string): React.CSSProperties {
  const color = STATUS_COLORS[status?.toLowerCase()] || '#605e5c';
  return {
    display: 'inline-block',
    padding: '2px 10px',
    borderRadius: 12,
    background: color,
    color: '#fff',
    fontWeight: 600,
    fontSize: 12,
    whiteSpace: 'nowrap',
  };
}

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

function toServerRelativePath(url: string, siteUrl: string): string {
  if (!url) return '';
  if (url.indexOf('/') === 0 && url.indexOf('//') !== 0) return url;
  try {
    const parsed = new URL(url);
    return parsed.pathname;
  } catch {
    try {
      const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
      return `${sitePath}/${url.replace(/^\//, '')}`;
    } catch {
      return url;
    }
  }
}

function parseRowsFromArrayBuffer(buffer: ArrayBuffer): IExcelRow[] {
  const data = new Uint8Array(buffer);
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  // ✅ First attempt: normal structured parsing
  const json: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (json && json.length > 0) {

    const normalizeKey = (key: string): string =>
      key.replace(/\s+/g, '').toLowerCase();

    const result: IExcelRow[] = [];

    for (let i = 0; i < json.length; i++) {
      const row = json[i];
      const normalizedRow: Record<string, any> = {};

      for (const key in row) {
        if (row.hasOwnProperty(key)) {
          normalizedRow[normalizeKey(key)] = row[key];
        }
      }

      const item: IExcelRow = {
        SNo: String(normalizedRow['sno'] || normalizedRow['serialno'] || ''),
        Particulars: String(normalizedRow['particulars'] || normalizedRow['description'] || ''),
        Department: String(normalizedRow['department'] || ''),
        DocumentLink: String(normalizedRow['documentlink'] || ''),
        Member: String(normalizedRow['member'] || ''),
        Status: String(normalizedRow['status'] || ''),
        Remarks: String(normalizedRow['remarks'] || ''),
        CompliedAsOn: String(normalizedRow['compliedason'] || normalizedRow['date'] || '')
      };

      // ✅ Remove empty rows safely
      let hasValue = false;
      for (const key in item) {
        if (item.hasOwnProperty(key)) {
          if (String(item[key as keyof IExcelRow] || '').trim() !== '') {
            hasValue = true;
            break;
          }
        }
      }

      if (hasValue) result.push(item);
    }

    return result;
  }

  // ✅ FALLBACK for messy Excel
  const rowsArr: any[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ''
  });

  if (!rowsArr || rowsArr.length === 0) return [];

  // ✅ find header row safely (NO findIndex)
  let headerIndex = -1;

  for (let i = 0; i < rowsArr.length; i++) {
    const rowStr = rowsArr[i].join(' ').toLowerCase();
    if (rowStr.indexOf('particulars') !== -1) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) headerIndex = 0;

  const dataRows = rowsArr.slice(headerIndex + 1);

  const finalRows: IExcelRow[] = [];

  let lastSNo = '';
  let lastParticular = '';

  for (let i = 0; i < dataRows.length; i++) {

    const r = dataRows[i];
    const cells: string[] = [];

    for (let j = 0; j < r.length; j++) {
      cells.push(String(r[j] || '').trim());
    }

    let SNo = cells[0];
    let Particulars = cells[1];
    let Department = cells[2];
    let Member = cells[3];
    let Status = '';
    let Remarks = cells[5];
    let CompliedAsOn = cells[6];

    // ✅ Handle sub rows
    if (!SNo || /^[a-z]\)/i.test(SNo)) {
      SNo = lastSNo;
    } else {
      lastSNo = SNo;
    }

    // ✅ Handle section headers
    if (Particulars) {
      lastParticular = Particulars;
    } else {
      Particulars = lastParticular;
    }

    // ✅ Safe status detection (NO includes/find)
    for (let j = 0; j < cells.length; j++) {
      const val = String(cells[j] || '').toLowerCase();

      if (
        val === 'complied' ||
        val === 'pending' ||
        val === 'not applicable' ||
        val === 'deferred' ||
        val === 'completed' ||
        val === 'rejected'
      ) {
        Status = cells[j];
        break;
      }
    }

    const item: IExcelRow = {
      SNo: SNo || '',
      Particulars: Particulars || '',
      Department: Department || '',
      DocumentLink: '',
      Member: Member || '',
      Status: Status || '',
      Remarks: Remarks || '',
      CompliedAsOn: CompliedAsOn || ''
    };

    // ✅ Remove empty row
    let hasValue = false;
    for (const key in item) {
      if (item.hasOwnProperty(key)) {
        if (String(item[key as keyof IExcelRow] || '').trim() !== '') {
          hasValue = true;
          break;
        }
      }
    }

    if (hasValue) {
      finalRows.push(item);
    }
  }

  return finalRows;
}
export default class FileUploaed extends React.Component<IFileUploaedProps, IFileUploaedState> {
  constructor(props: IFileUploaedProps) {
    super(props);
    this.state = {
      rows: [],
      configMap: {},
      configLoading: true,
      configError: '',
      error: '',
      fileName: '',
      dataSource: '',
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
    this._loadSavedExcel();
  }

  // ── Load ConfigurationList ──────────────────────────────────
  private _loadConfigurationList(): void {
    const { spHttpClient, siteUrl } = this.props;
    const url = `${siteUrl}/_api/web/lists/getbytitle('ConfigurationList')/items?$select=MappingField,LocationURL&$top=5000`;

    spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: { MappingField: string; LocationURL: string }[] }) => {
        const configMap: Record<string, string> = {};
        (data.value || []).forEach(item => {
          if (item.MappingField) {
            configMap[String(item.MappingField).trim()] = item.LocationURL || '';
          }
        });
        this.setState({ configMap, configLoading: false });
      })
      .catch(() => {
        this.setState({ configError: 'Could not load ConfigurationList from SharePoint.', configLoading: false });
      });
  }

  // ── Load saved Excel from ComplianceBulkUploads on startup ──
  private _loadSavedExcel(): void {
    const { spHttpClient, siteUrl } = this.props;
    const sitePath = toServerRelativePath(siteUrl, siteUrl);
    const fileUrl = `${sitePath}/${SAVED_LIBRARY}/${SAVED_FILE_NAME}`;
    const apiUrl  = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')/$value`;
    console.log('Checking for saved Excel at:', apiUrl);
    spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => {
        if (!res.ok) return; // file doesn't exist yet — silent, no error shown
        return res.arrayBuffer().then((buffer: ArrayBuffer) => {
          try {
            const rows = parseRowsFromArrayBuffer(buffer);
            if (rows.length > 0) {
              this.setState({ rows, fileName: SAVED_FILE_NAME, dataSource: 'sharepoint' });
            }
          } catch {
            // corrupted saved file — ignore silently
          }
        });
      })
      .catch(() => { /* library may not exist yet — ignore */ });
  }

  // ── Save / overwrite Excel in ComplianceBulkUploads ─────────
  private _saveExcelToLibrary(arrayBuffer: ArrayBuffer): void {
    const { spHttpClient, siteUrl } = this.props;
    const sitePath  = toServerRelativePath(siteUrl, siteUrl);
    const folderUrl = `${sitePath}/${SAVED_LIBRARY}`;
    const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderUrl)}')/Files/add(overwrite=true,url='${encodeURIComponent(SAVED_FILE_NAME)}')`;

    this.setState({ savingExcel: true, saveError: '' });

    spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' },
      body: arrayBuffer as unknown as string,
    })
    .then((res: SPHttpClientResponse) => {
      if (res.ok) {
        this.setState({ savingExcel: false });
      } else {
        return res.json().then((err: { error?: { message?: { value?: string } } }) => {
          const msg = (err.error && err.error.message && err.error.message.value) || 'Could not save file to SharePoint.';
          this.setState({ savingExcel: false, saveError: msg });
        });
      }
    })
    .catch(() => {
      this.setState({ savingExcel: false, saveError: 'Could not save file to SharePoint library.' });
    });
  }

  // ── Handle local Excel import ────────────────────────────────
  private handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt): void => {
      const buffer = (evt.target as FileReader).result as ArrayBuffer;
      try {
        const rows = parseRowsFromArrayBuffer(buffer);
        this.setState({ rows, error: '', fileName: file.name, dataSource: 'local' });
        // Save to ComplianceBulkUploads, overwriting any previous version
        this._saveExcelToLibrary(buffer);
      } catch {
        this.setState({ error: 'Failed to parse Excel file. Please check the file format.', rows: [], fileName: '' });
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // ── Modal helpers ────────────────────────────────────────────
  private _openFolderModal = (SNo: string, folderUrl: string): void => {
    this.setState({
      modalOpen: true,
      modalTitle: `Files for S.No ${SNo}`,
      modalFolderUrl: folderUrl,
      modalFiles: [],
      modalLoading: true,
      modalError: '',
    }, () => this._loadFolderFiles(folderUrl));
  };

  private _loadFolderFiles(folderUrl: string): void {
    const { spHttpClient, siteUrl } = this.props;
    const serverRelPath = toServerRelativePath(folderUrl, siteUrl);
    const encodedPath   = encodeURIComponent(serverRelPath);
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$orderby=Name`;

    spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: ISharePointFile[] }) => {
        if (data.value) {
          this.setState({ modalFiles: data.value, modalLoading: false });
        } else {
          this.setState({ modalError: 'No files found or folder does not exist.', modalLoading: false });
        }
      })
      .catch(() => {
        this.setState({ modalError: 'Failed to load files from SharePoint folder.', modalLoading: false });
      });
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
    const serverRelPath = toServerRelativePath(modalFolderUrl, siteUrl);
    const encodedPath   = encodeURIComponent(serverRelPath);
    const encodedName   = encodeURIComponent(file.name);
    const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Files/add(overwrite=true,url='${encodedName}')`;

    this.setState({ uploadingFile: true, uploadError: '', uploadSuccess: '' });

    const reader = new FileReader();
    reader.onload = (evt): void => {
      const arrayBuffer = (evt.target as FileReader).result as ArrayBuffer;
      spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' },
        body: arrayBuffer as unknown as string,
      })
      .then((res: SPHttpClientResponse) => {
        if (res.ok) {
          this.setState({ uploadingFile: false, uploadSuccess: `"${file.name}" uploaded successfully.` });
          this._loadFolderFiles(modalFolderUrl);
        } else {
          return res.json().then((err: { error?: { message?: { value?: string } } }) => {
            const msg = (err.error && err.error.message && err.error.message.value) || 'Upload failed.';
            this.setState({ uploadingFile: false, uploadError: msg });
          });
        }
      })
      .catch(() => {
        this.setState({ uploadingFile: false, uploadError: 'Upload failed. Please try again.' });
      });
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Modal render ─────────────────────────────────────────────
  private _renderModal(): React.ReactElement | null {
    const { modalOpen, modalTitle, modalFiles, modalLoading, modalError, uploadingFile, uploadError, uploadSuccess } = this.state;
    if (!modalOpen) return null;

    return (
      <div className={styles.modalOverlay} onClick={this._closeModal}>
        <div className={styles.modalBox} onClick={e => e.stopPropagation()}>

          <div className={styles.modalHeader}>
            <div className={styles.modalHeaderLeft}>
              <div className={styles.modalFolderIcon}>📁</div>
              <div>
                <div className={styles.modalTitle}>{modalTitle}</div>
              </div>
            </div>
            <button className={styles.modalClose} onClick={this._closeModal} title="Close">✕</button>
          </div>

          <div className={styles.modalBody}>
            <div className={styles.uploadSection}>
              <label className={styles.uploadSectionLabel} htmlFor="modalFileUpload">
                {uploadingFile
                  ? <><div className={styles.spinnerSm} /> Uploading...</>
                  : <>⬆ Upload File to this Folder</>}
              </label>
              <input
                id="modalFileUpload"
                type="file"
                className={styles.fileInput}
                onChange={this._handleUpload}
                disabled={uploadingFile}
              />
            </div>
            {uploadError   && <div className={styles.errorMsg}>{uploadError}</div>}
            {uploadSuccess && <div className={styles.successMsg}>{uploadSuccess}</div>}

            <div className={styles.modalDivider} />

            {modalLoading && (
              <div className={styles.modalSpinner}>
                <div className={styles.spinner} />
                <span>Loading files...</span>
              </div>
            )}
            {!modalLoading && modalError && <div className={styles.errorMsg}>{modalError}</div>}
            {!modalLoading && !modalError && modalFiles.length === 0 && (
              <div className={styles.modalEmpty}>No files found in this folder.</div>
            )}
            {!modalLoading && !modalError && modalFiles.length > 0 && (
              <>
                <div className={styles.modalFileCount}>{modalFiles.length} file{modalFiles.length !== 1 ? 's' : ''} in folder</div>
                <table className={styles.modalTable}>
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>File Name</th>
                      <th>Size</th>
                      <th>Last Modified</th>
                      <th>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {modalFiles.map((f, idx) => {
                      const absUrl = this._toAbsoluteUrl(f.ServerRelativeUrl);
                      const viewUrl = this._viewerUrl(f.Name, absUrl);
                      return (
                        <tr key={idx} className={idx % 2 === 0 ? styles.rowEven : styles.rowOdd}>
                          <td className={styles.indexCell}>{idx + 1}</td>
                          <td className={styles.fileNameCell}>
                            <span className={styles.fileIcon}>{this._fileIcon(f.Name)}</span>
                            {f.Name}
                          </td>
                          <td>{formatBytes(f.Length)}</td>
                          <td>{formatDate(f.TimeLastModified)}</td>
                          <td>
                            <div className={styles.actionBtns}>
                              <a href={viewUrl} target="_blank" rel="noreferrer" className={styles.btnView} title="View file in browser">
                                👁 View
                              </a>
                              <a href={absUrl} download={f.Name} className={styles.btnDownload} title="Download file">
                                ⬇ Download
                              </a>
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
    if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].indexOf(ext) !== -1) {
      return absUrl + '?web=1';
    }
    return absUrl;
  }

  private _toAbsoluteUrl(serverRelativeUrl: string): string {
    try {
      const origin = new URL(this.props.siteUrl).origin;
      return origin + serverRelativeUrl;
    } catch {
      return serverRelativeUrl;
    }
  }

  private _fileIcon(name: string): string {
    const ext = name.split('.').pop()?.toLowerCase() || '';
    const map: Record<string, string> = {
      pdf: '📕', doc: '📘', docx: '📘', xls: '📗', xlsx: '📗',
      ppt: '📙', pptx: '📙', png: '🖼', jpg: '🖼', jpeg: '🖼',
      gif: '🖼', zip: '🗜', txt: '📄', csv: '📊',
    };
    return map[ext] || '📄';
  }

  // ── Main render ──────────────────────────────────────────────
  public render(): React.ReactElement<IFileUploaedProps> {
    const { rows, configMap, configLoading, configError, error, fileName, dataSource, savingExcel, saveError } = this.state;

    return (
      <section className={styles.fileUploaed}>

        {/* ── Branded Header ── */}
        <div className={styles.header}>
          <div className={styles.headerInner}>
            <div className={styles.headerLeft}>
              <img
                src={require('../assets/dmi-logo.svg')}
                alt="DMI Finance"
                className={styles.brandLogo}
              />
              <div className={styles.headerDivider} />
              <div>
                <h2 className={styles.title}>Bulk Upload Tracker</h2>
                <p className={styles.subtitle}>Enterprise Document Management System</p>
              </div>
            </div>
            <span className={styles.headerBadge}>Internal Tool</span>
          </div>
        </div>

        {/* ── Status messages ── */}
        {configLoading && (
          <div className={styles.infoMsg}>
            <span>⏳</span> Loading configuration from SharePoint...
          </div>
        )}
        {configError  && <div className={styles.errorMsg}>{configError}</div>}
        {saveError    && <div className={styles.errorMsg}>⚠ Could not save to library: {saveError}</div>}

        {/* ── Toolbar ── */}
        <div className={styles.toolbar}>
          <div className={styles.toolbarLeft}>
            <label className={styles.uploadLabel} htmlFor="excelUpload">
              <span className={styles.uploadIcon}>📊</span>
              <span>{savingExcel ? 'Saving...' : 'Import Excel File'}</span>
            </label>
            <input
              id="excelUpload"
              type="file"
              multiple
              accept=".xlsx,.xls"
              onChange={this.handleFileChange}
              className={styles.fileInput}
              disabled={savingExcel}
            />
            {dataSource === 'sharepoint' && (
              <span className={styles.sourceChip} title={`Auto-loaded: ${SAVED_LIBRARY}/${SAVED_FILE_NAME}`}>
                ☁ Loaded from SharePoint
              </span>
            )}
            {dataSource === 'local' && fileName && (
              <span className={styles.sourceChipLocal} title={`Saved to ${SAVED_LIBRARY}/${SAVED_FILE_NAME}`}>
                {savingExcel ? '⏳ Saving to SharePoint...' : `✔ Saved as ${SAVED_FILE_NAME}`}
              </span>
            )}
          </div>
          {rows.length > 0 && (
            <span className={styles.recordBadge}>
              <span className={styles.recordDot} />
              {rows.length} Record{rows.length !== 1 ? 's' : ''} Loaded
            </span>
          )}
        </div>

        {error && <div className={styles.errorMsg}>{error}</div>}

        {/* ── Data Table ── */}
        {rows.length > 0 && (
          <div className={styles.tableWrapper}>
            <div className={styles.tableInfo}>
              Showing all {rows.length} record{rows.length !== 1 ? 's' : ''} — click &ldquo;View Files&rdquo; to browse SharePoint documents
            </div>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>S.No</th>
                  <th>Particulars</th>
                  <th>Department</th>
                  <th>Document Link</th>
                  <th>Member</th>
                  <th>Status</th>
                  <th>Remarks</th>
                  <th>Complied As On</th>
                  <th>SharePoint Files</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, i) => {
                  const spUrl = configMap[row.SNo.trim()] || '';
                  return (
                    <tr key={i} className={i % 2 === 0 ? styles.rowEven : styles.rowOdd}>
                      <td className={styles.indexCell}>{row.SNo}</td>
                      <td>{row.Particulars}</td>
                      <td>{row.Department}</td>
                      <td>
                        {row.DocumentLink
                          ? <a href={row.DocumentLink} target="_blank" rel="noreferrer" className={styles.docLink}>↗ View</a>
                          : <span className={styles.noLink}>—</span>}
                      </td>
                      <td>{row.Member}</td>
                      <td>
                        {row.Status
                          ? <span style={statusBadgeStyle(row.Status)}>{row.Status}</span>
                          : '—'}
                      </td>
                      <td>{row.Remarks || <span className={styles.noLink}>—</span>}</td>
                      <td>{row.CompliedAsOn || <span className={styles.noLink}>—</span>}</td>
                      <td>
                        {spUrl
                          ? (
                            <button
                              className={styles.spLinkBtn}
                              onClick={() => this._openFolderModal(row.SNo, spUrl)}
                            >
                              📁 View Files
                            </button>
                          )
                          : <span className={styles.noLink}>—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {rows.length === 0 && !error && (
          <div className={styles.emptyState}>
            <div className={styles.emptyIconWrap}>📊</div>
            <p className={styles.emptyTitle}>No Records Loaded</p>
            <p className={styles.emptyDesc}>
              Import an Excel file using the button above.<br />
              Expected columns: <strong>S.No, Particulars, Department, Document Link, Member, Status</strong>
            </p>
          </div>
        )}

        {this._renderModal()}
      </section>
    );
  }
}
