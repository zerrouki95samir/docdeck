import type { DocKind } from '../lenses';

/* ---------- icons ---------- */

function IconSheet({ size = 22 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <rect x="3" y="3" width="18" height="18" rx="2.5" stroke="currentColor" strokeWidth="1.6" />
      <path d="M3 9.5h18M3 15h18M9.5 3v18" stroke="currentColor" strokeWidth="1.6" />
    </svg>
  );
}

function IconDoc({ size = 22 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path d="M6 2.75h7.5L19 8.25v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-14.5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.6" />
      <path d="M13 3v5.5h6" stroke="currentColor" strokeWidth="1.6" />
      <path d="M7.5 12.5h7M7.5 16h5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  );
}

function IconPdf({ size = 22 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path d="M6 2.75h7.5L19 8.25v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-14.5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.6" />
      <path d="M13 3v5.5h6" stroke="currentColor" strokeWidth="1.6" />
      <path d="M7.4 17.2V11h1.9a2.1 2.1 0 0 1 0 4.2H7.4m8.6 2v-6.4H18" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  );
}

function IconFolder({ size = 40 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path d="M3.5 6.5a2 2 0 0 1 2-2h4l2 2.5h6.5a1.6 1.6 0 0 1 1.6 1.6v9.4a2 2 0 0 1-2 2H5.5a2 2 0 0 1-2-2v-11a.5.5 0 0 1 .06-.24Z" fill="currentColor" opacity=".35" />
      <path d="M3.5 6.5a2 2 0 0 1 2-2h4l2 2.5h6.5a1.6 1.6 0 0 1 1.6 1.6v9.4a2 2 0 0 1-2 2H5.5a2 2 0 0 1-2-2v-11a.5.5 0 0 1 .06-.24Z" stroke="currentColor" strokeWidth="1.6" />
    </svg>
  );
}

function IconPlus({ size = 16 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path d="M12 5v14M5 12h14" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
    </svg>
  );
}

function LogoMark({ size = 46 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 512 512" fill="none" aria-hidden>
      <defs>
        <linearGradient id="dd-logo-bg" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#17395F" />
          <stop offset="1" stopColor="#0A1830" />
        </linearGradient>
        <clipPath id="dd-logo-gt">
          <rect x="124" y="224" width="126" height="172" rx="12" />
        </clipPath>
      </defs>
      <rect width="512" height="512" rx="112" fill="url(#dd-logo-bg)" />
      <path d="M116 196 v-58 c0-11 9-20 20-20 h74 c11 0 20 9 20 20 v58 Z" fill="#22A75D" />
      <path d="M219 196 v-58 c0-11 9-20 20-20 h74 c11 0 20 9 20 20 v58 Z" fill="#2F6FED" />
      <path d="M322 196 v-58 c0-11 9-20 20-20 h34 c14 0 26 9 26 20 v58 Z" fill="#E5484D" />
      <g fill="#FFFFFF" opacity="0.92">
        <rect x="141" y="137" width="24" height="22" rx="4" />
        <rect x="173" y="137" width="24" height="22" rx="4" />
        <rect x="141" y="167" width="24" height="22" rx="4" />
        <rect x="173" y="167" width="24" height="22" rx="4" />
      </g>
      <g fill="#FFFFFF" opacity="0.92">
        <rect x="243" y="135" width="66" height="13" rx="6.5" />
        <rect x="243" y="158" width="66" height="13" rx="6.5" />
        <rect x="243" y="181" width="44" height="11" rx="5.5" />
      </g>
      <g>
        <rect x="364" y="130" width="60" height="56" rx="5" fill="#FFFFFF" opacity="0.92" />
        <path d="M402 130 h22 v22 Z" fill="#C33B40" />
      </g>
      <rect x="96" y="196" width="320" height="228" rx="22" fill="#F8FAFF" />
      <g clipPath="url(#dd-logo-gt)">
        <rect x="124" y="224" width="126" height="172" fill="#DFF2E6" />
        <rect x="124" y="224" width="126" height="38" fill="#79C89B" />
        <path d="M166 262 h7 v134 h-7 Z M207 262 h7 v134 h-7 Z" fill="#F8FAFF" />
        <path d="M124 281 h126 v9 H124 Z M124 338 h126 v9 H124 Z" fill="#F8FAFF" />
        <g fill="#22A75D" opacity="0.6">
          <rect x="133" y="290" width="27" height="39" rx="3" />
          <rect x="174" y="290" width="27" height="39" rx="3" />
          <rect x="215" y="347" width="27" height="39" rx="3" />
        </g>
      </g>
      <rect x="278" y="232" width="110" height="16" rx="8" fill="#2F6FED" opacity="0.85" />
      <rect x="278" y="262" width="110" height="12" rx="6" fill="#B9CDFC" />
      <rect x="278" y="284" width="82" height="12" rx="6" fill="#CBD6E4" />
      <rect x="278" y="312" width="110" height="12" rx="6" fill="#CBD6E4" />
      <rect x="278" y="334" width="94" height="12" rx="6" fill="#CBD6E4" />
      <rect x="278" y="358" width="58" height="36" rx="10" fill="#FBE3E4" opacity="0.9" />
      <rect x="286" y="368" width="42" height="6" rx="3" fill="#E5484D" />
      <rect x="286" y="380" width="30" height="6" rx="3" fill="#F0A9AC" />
    </svg>
  );
}

/* ---------- kinds ---------- */

const KIND_META: Record<
  DocKind,
  { label: string; formats: string[]; blurb: string; bullets: string[]; tone: string }
> = {
  sheet: {
    label: 'Spreadsheet',
    formats: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'],
    blurb: 'Formula engine, cell styling, multi-sheet workbooks.',
    bullets: ['Formulas, ranges & references', 'Fonts, formats, merged cells', 'Sort, filter & export'],
    tone: 'cd-green',
  },
  docx: {
    label: 'Word Document',
    formats: ['docx'],
    blurb: 'Writer with live preview, styles and tables.',
    bullets: ['Headings & rich character styles', 'Tables, images, hyperlinks', 'True .docx save, ready for Word'],
    tone: 'cd-blue',
  },
  pdf: {
    label: 'PDF',
    formats: ['pdf'],
    blurb: 'Reader with annotation and page tooling.',
    bullets: ['Zoom, search, thumbnails', 'Highlight & freehand markup', 'Rotate, reorder, merge, forms'],
    tone: 'cd-red',
  },
};

const extKindsFor = (k: DocKind) => (k === 'sheet' ? [...KIND_META.sheet.formats, 'uni'] : KIND_META[k].formats);

export interface HomeProps {
  onCreate: (kind: DocKind) => void;
  onOpen: () => void;
  onOpenPath: (path: string) => void;
  recents: string[];
  resume: Array<{ kind: DocKind; name: string; dirty: boolean }>;
  onResume: (kind: DocKind) => void;
}

export function Home({ onCreate, onOpen, onOpenPath, recents, resume, onResume }: HomeProps) {
  const kindIcon = (k: DocKind) =>
    k === 'pdf' ? <IconPdf /> : k === 'docx' ? <IconDoc /> : <IconSheet size={20} />;

  return (
    <div className="home">
      <div className="home-inner">
        {/* hero */}
        <header className="home-hero">
          <div className="home-logo-row">
            <LogoMark />
            <div>
              <h1 className="home-title">DocDeck</h1>
              <p className="home-sub">Spreadsheets, Word documents and PDFs — one tabbed deck for all of them.</p>
            </div>
          </div>
          <div className="hero-actions">
            <button className="home-open-btn" onClick={onOpen} title="Open a workbook, document or PDF">
              <IconFolder size={18} />
              Open a File…
            </button>
            <span className="home-drop-hint">…or drop files anywhere in this window</span>
          </div>
        </header>

        {/* create */}
        <section className="home-section">
          <h2 className="home-h2">Create a new file</h2>
          <div className="home-cards">
            {(['sheet', 'docx', 'pdf'] as const).map((k) => {
              const m = KIND_META[k];
              return (
                <article className="home-card" key={k}>
                  <div className={`home-card-icon ${m.tone}`}>{kindIcon(k)}</div>
                  <h3 className="home-card-title">{m.label}</h3>
                  <p className="home-card-blurb">{m.blurb}</p>
                  <ul className="home-card-list">
                    {m.bullets.map((b) => (
                      <li key={b}>{b}</li>
                    ))}
                  </ul>
                  <div className="home-card-formats">
                    {extKindsFor(k).map((f) => (
                      <span className="format-chip" key={f}>
                        .{f}
                      </span>
                    ))}
                  </div>
                  <button className="home-create-btn" onClick={() => onCreate(k)}>
                    <IconPlus size={15} />
                    Create {m.label}
                  </button>
                </article>
              );
            })}
          </div>
        </section>

        {/* resume open docs */}
        {resume.length > 0 && (
          <section className="home-section">
            <h2 className="home-h2">Still open in this window</h2>
            <div className="home-resume">
              {resume.map((r) => (
                <button className="resume-chip" key={r.kind} onClick={() => onResume(r.kind)} title="Switch to this document">
                  {kindIcon(r.kind)}
                  <span className="resume-name">{r.name}</span>
                  {r.dirty && <span className="resume-dirty">unsaved</span>}
                </button>
              ))}
            </div>
          </section>
        )}

        {/* recents */}
        <section className="home-section">
          <h2 className="home-h2">Recent files</h2>
          {recents.length === 0 ? (
            <p className="home-empty">Files you open will appear here for one-click access.</p>
          ) : (
            <ul className="home-recents">
              {recents.map((p) => {
                const lower = p.toLowerCase();
                const k: DocKind | null = /\.(xlsx|xlsm|xls|csv|tsv|uni)$/.test(lower)
                  ? 'sheet'
                  : /\.pdf$/.test(lower)
                    ? 'pdf'
                    : /\.docx$/.test(lower)
                      ? 'docx'
                      : null;
                return (
                  <li key={p}>
                    <button className="recent-row" onClick={() => onOpenPath(p)} title={p}>
                      {k ? (
                        kindIcon(k)
                      ) : (
                        <IconFolder size={18} />
                      )}
                      <span className="recent-name">{p.split('/').pop()}</span>
                      <span className="recent-path">{p}</span>
                    </button>
                  </li>
                );
              })}
            </ul>
          )}
        </section>

        <footer className="home-foot">
          Spreadsheets · Word · PDF — everything DocDeck lays out stays intact: styles, formulas,
          annotations and structure round-trip cleanly.
        </footer>
      </div>
    </div>
  );
}

export default Home;
